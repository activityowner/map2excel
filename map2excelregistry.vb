Imports Microsoft.Win32
Module map2excelregistry

    Sub setmtckey(ByVal folder As String, ByVal key As String, ByVal keyvalue As String)
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.OpenSubKey("Software", True)
        regKey = regKey.CreateSubKey("ActivityOwner.Com", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
        regKey = regKey.CreateSubKey("map2excel", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
        regKey = regKey.CreateSubKey(folder, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
        regKey.SetValue(key, keyvalue)
        regKey = Nothing
    End Sub
    Function getmtckey(ByVal folder As String, ByVal key As String) As String
        Dim regKey As RegistryKey
        Try
            regKey = Registry.CurrentUser.OpenSubKey("Software", Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree)
            regKey = regKey.OpenSubKey("ActivityOwner.Com", Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree)
            regKey = regKey.OpenSubKey("map2excel", Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree)
            regKey = regKey.OpenSubKey(folder, Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree)
            getmtckey = CType(regKey.GetValue(key), String)
        Catch
            getmtckey = ""
        End Try
    End Function

End Module
