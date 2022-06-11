Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Module module2
    'ao_map2outline 29Dec2009 http://creativecommons.org/licenses/by-nc-nd/3.0/
    '#uses "ao_common.mmbas"
    Sub map2outline(ByRef m_app As mm.Application)
        Dim parent As mm.Topic
        Dim t As mm.Topic
        Dim f As String
        Dim indent As Integer
        indent = 0
        If m_app.ActiveDocument.Selection.Count > 0 Then
            parent = m_app.ActiveDocument.Selection.PrimaryTopic
        Else
            parent = m_app.ActiveDocument.CentralTopic
        End If
        If parent.Document.IsModified Then parent.Document.Save()
        f = Replace(parent.Document.FullName, ".mmap", ".html")
        Microsoft.VisualBasic.FileOpen(1, f, OpenMode.Output, OpenAccess.Write)
        Microsoft.VisualBasic.Print(1, "<html>")
        Microsoft.VisualBasic.Print(1, parent.Text)
        For Each t In parent.AllSubTopics
            exportinfo(t, indent)
        Next
        Microsoft.VisualBasic.Print(1, "</html>")
        Microsoft.VisualBasic.FileClose(1)
        On Error Resume Next
        Shell("C:\Program Files\Internet Explorer\iexplore.exe " & f, Microsoft.VisualBasic.AppWinStyle.NormalFocus)
        parent = Nothing
        t = Nothing
    End Sub
  
    Sub exportinfo(ByRef t As mm.Topic, ByVal indent As Integer)
        Dim st As mm.Topic
        Dim i As Integer
        If indent > 0 Then
            For i = 1 To indent
                Microsoft.VisualBasic.Print(1, "&nbsp;")
            Next
        Else
            Microsoft.VisualBasic.Print(1, "<hr>")
        End If
        Microsoft.VisualBasic.Print(1, hyperlinkstring(t) & "<br>" & vbCrLf)
        indent = indent + 1
        For Each st In t.AllSubTopics
            exportinfo(st, indent)
        Next
        st = Nothing
    End Sub
End Module
