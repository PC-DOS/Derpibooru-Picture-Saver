Imports System.ComponentModel
Imports Microsoft.Win32
Module WebBrowserVersionControl
    Public Sub SetWebBrowserFeatures(Version As Integer)
        If LicenseManager.UsageMode <> LicenseUsageMode.Runtime Then
            Exit Sub
        End If
        Dim ApplicationName = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName)
        Dim IEMode As UInt32 = GeoEmulationMode(Version)
        Dim FeatureControlRegistryKey = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\"
        Registry.SetValue(FeatureControlRegistryKey & "FEATURE_BROWSER_EMULATION", ApplicationName, IEMode, RegistryValueKind.DWord)
        Registry.SetValue(FeatureControlRegistryKey + "FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION", ApplicationName, 1, RegistryValueKind.DWord)
    End Sub
    Public Function GetWebBrowserVersion() As Integer
        Dim BrowserVersion As Integer = 0
        Dim IERegistry = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Internet Explorer", RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.QueryValues)
        Dim Version = IERegistry.GetValue("svcVersion")
        If IsNothing(Version) Then
            Version = IERegistry.GetValue("Version")
            If IsNothing(Version) Then
                Return -1
            End If
        End If
        BrowserVersion = Version.ToString.Split(".")(0)
        GetWebBrowserVersion = BrowserVersion
    End Function
    Public Function GeoEmulationMode(Version As Integer) As UInt32
        Dim IEMode As UInt32 = 11000
        Select Case Version
            Case 7
                IEMode = 7000
            Case 8
                IEMode = 8000
            Case 9
                IEMode = 9000
            Case 10
                IEMode = 10000
            Case 11
                IEMode = 11000
        End Select
        Return IEMode
    End Function
End Module
