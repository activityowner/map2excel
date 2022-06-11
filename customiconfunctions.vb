Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Module customiconfunctions
    Public sigsiglist(100) As String
    Public sigfilelist(100) As String
    Public sigcount As Integer = 0
    Function FindCustomIconFileNameFromSignature(ByRef m_app As mm.Application, ByVal signature As String, ByVal icondir As String) As String
        'Debug.Print("in findcustomiconfilenamefromsignature")
        If Len(icondir) = 0 Then
            icondir = m_app.GetPath(mm.MmDirectory.mmDirectoryIcons)
        End If
        Dim di As New IO.DirectoryInfo(icondir)
        Dim diar1 As IO.FileInfo() = di.GetFiles()
        Dim dra As IO.FileInfo
        FindCustomIconFileNameFromSignature = ""
        'list the names of all files in the specified directory
        Dim i As Integer
        If sigcount > 0 Then
            For i = 1 To sigcount
                'Debug.Print(sigfilelist(i))
                If StrComp(signature, sigsiglist(i)) = 0 Then
                    'Debug.Print("found it in history" & sigfilelist(i))
                    FindCustomIconFileNameFromSignature = sigfilelist(i)
                    Exit Function
                End If
            Next
        End If
        For Each dra In diar1
            If StrComp(m_app.Utilities.GetCustomIconSignature(dra.FullName), signature) = 0 Then
                FindCustomIconFileNameFromSignature = dra.FullName
                'Debug.Print("found match for " & dra.FullName)
                sigcount = sigcount + 1
                sigsiglist(sigcount) = signature
                sigfilelist(sigcount) = dra.FullName
                Exit Function
            End If
        Next

        Dim diar2 As IO.DirectoryInfo() = di.GetDirectories()
        Dim dra2 As IO.DirectoryInfo
        Dim output As String
        For Each dra2 In diar2
            output = FindCustomIconFileNameFromSignature(m_app, signature, dra2.FullName)
            If Len(output) > 0 Then
                FindCustomIconFileNameFromSignature = output
                'Debug.Print("found match for " & dra2.FullName)
                sigcount = sigcount + 1
                sigsiglist(sigcount) = signature
                sigfilelist(sigcount) = dra2.FullName
                'Debug.Print(sigfilelist(sigcount))
                Exit Function
            End If
        Next
    End Function
End Module
