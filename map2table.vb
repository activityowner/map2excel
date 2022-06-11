Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Module module3
    Sub map2table(ByVal m_app As mm.application)
        Dim mains(150) As String 'assume <150 unique strings in each layer
        Dim subs(150) As String
        Dim temp(150) As String 'for copying over
        Dim swap As String
        Dim entries(1000) As mm.Topic 'assume <1000 third layer entries
        Dim maintext(1000) As Integer
        Dim subtext(1000) As Integer
        Dim mcount As Integer
        Dim scount As Integer
        Dim ecount As Integer
        Dim tempcount As Integer
        Dim f As String
        Dim mindex As Integer
        Dim sindex As Integer
        Dim autoopen As Boolean
        Dim include_link As Boolean
        Dim use_this_link As Boolean
        Dim include_notes As Boolean
        Dim bullet As Boolean
        Dim mark_no_children As Boolean
        Dim shorten_entries As Boolean
        Dim found As Boolean
        Dim mfound As Boolean
        Dim sfound As Boolean
        Dim first As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim addr As String
        Dim txt As String
        Dim imark As String
        Dim omark As String
        Dim tasks As String
        Dim main_as_row As Boolean
        Dim parenttopic As mm.Topic
        Dim mtopic As mm.Topic
        Dim stopic As mm.Topic
        Dim sstopic As mm.Topic
        Dim tasktopic As mm.Topic '4th layer items
        Dim separator As String
        Dim max_length As Integer
        Dim answer As Integer
        Dim doc As mm.Document
        Dim usedefaults As Boolean
        Dim guid As String
        Dim prefix As String
        Dim postfix As String

        max_length = 250
        '--------------------------------------------------------------------------------------------------
        mcount = 0                'initialize unique string counts
        scount = 0                '
        ecount = 0
        usedefaults = Command() = "default"
        doc = m_app.ActiveDocument  'lock onto active document
        If doc.IsModified Then
            doc.Save()
            'MsgBox("Save document before running map2table") 'make sure we have an html name/destination
            Exit Sub
        End If
        '-------------Decide parent to work from----------------------------------------------
        parenttopic = Nothing
        If Not doc.Selection.PrimaryTopic Is Nothing Then
            If Not doc.Selection.PrimaryTopic.IsCentralTopic Then
                answer = MsgBox("Run on Branch" & doc.Selection.PrimaryTopic.Text & "?", vbYesNoCancel)
                If answer = vbYes Then
                    parenttopic = doc.Selection.PrimaryTopic
                End If
                If answer = vbNo Then
                    parenttopic = doc.CentralTopic
                End If
                If answer = vbCancel Then
                    Exit Sub
                End If
            Else
                parenttopic = doc.CentralTopic
            End If
        Else
            parenttopic = doc.CentralTopic
        End If
        If usedefaults Then
            main_as_row = False
        Else
            main_as_row = MsgBox("Do you want Main topics listed as Rows?", vbYesNo) = vbYes
        End If

        '-----MAP2TABLE OPTIONS--------------------------------------------------------------------------
        autoopen = True   'auto open html file?
        If usedefaults Then
            include_link = True
            include_notes = False
        Else
            include_link = MsgBox("Include Hyperlinks?", vbYesNo) = vbYes   'include hyperlinks in html table
            include_notes = MsgBox("Include Notes?", vbYesNo) = vbYes 'add notes under entries
        End If
        mark_no_children = True  'add * if no children under entry (e.g. no next actions)
        separator = "<br>" 'Can separate table entries with "<br>", "<hr>", Or ","
        If Not include_notes Then
            shorten_entries = True  'shorten long entries to end with ... depending on number of columns present
            If Not usedefaults Then
                bullet = MsgBox("Include Bullets", vbYesNo) = vbYes      'bullets in table by setting to false
            Else
                bullet = False
            End If
        Else
            bullet = False
            shorten_entries = False
        End If


        '------------Make lists of 1st/2nd layer words and index 3rd layer----------------------
        For Each mtopic In parenttopic.AllSubTopics 'make list of unique 1st layer topics
            mfound = False
            For i = 1 To mcount
                If mtopic.Text = mains(i) Then
                    mfound = True
                    mindex = i
                    Exit For
                End If
            Next
            If Not mfound Then
                mcount = mcount + 1
                mindex = mcount
                mains(mcount) = mtopic.Text
            End If
            For Each stopic In mtopic.AllSubTopics 'make list of 2nd layer topics
                sfound = False
                sindex = 0
                For i = 1 To scount
                    If stopic.Text = subs(i) Then
                        sfound = True
                        sindex = i
                        Exit For
                    End If
                Next
                If Not sfound Then
                    scount = scount + 1
                    sindex = scount
                    subs(scount) = stopic.Text
                End If
                For Each sstopic In stopic.AllSubTopics
                    ecount = ecount + 1
                    entries(ecount) = sstopic
                    subtext(ecount) = sindex
                    maintext(ecount) = mindex
                Next
            Next
        Next


        For i = 1 To scount - 1 'sort rows
            For j = i + 1 To scount
                If subs(i) > subs(j) Then
                    swap = subs(i)
                    subs(i) = subs(j)
                    subs(j) = swap
                    For k = 1 To ecount
                        If subtext(k) = i Then
                            subtext(k) = j
                        ElseIf subtext(k) = j Then
                            subtext(k) = i
                        End If
                    Next
                End If
            Next
        Next

        If main_as_row Then
            'move main to temp
            tempcount = scount
            For i = 1 To scount
                temp(i) = subs(i)
            Next
            scount = mcount
            For i = 1 To mcount
                subs(i) = mains(i)
            Next
            mcount = tempcount
            For i = 1 To mcount
                mains(i) = temp(i)
            Next
            For i = 1 To ecount
                j = subtext(i)
                subtext(i) = maintext(i)
                maintext(i) = j
            Next
        End If

        '----------
        'decide whether to make topics with no subtopics show in italics
        'if no entry has children then don't bother marking everything
        found = False
        For Each mtopic In parenttopic.AllSubTopics
            For Each stopic In mtopic.AllSubTopics
                For Each sstopic In stopic.AllSubTopics
                    If sstopic.AllSubTopics.Count > 0 Then found = True
                Next
            Next
        Next
        If Not found Then mark_no_children = False
        '
        '--------Start outputting info -------------------------------------------------------------
        f = Replace(doc.FullName, ".mmap", ".html") 'save doc to same location as map
        Microsoft.VisualBasic.FileOpen(1, f, OpenMode.Output, OpenAccess.Write)
        Microsoft.VisualBasic.Print(1, "<html><body><Table border=1><small><small>")
        '----Print the Header Row--------------------------------------------
        Microsoft.VisualBasic.Print(1, "<tr><th></th>") 'print hearder row with main layer topics
        For i = 1 To mcount
            Microsoft.VisualBasic.Print(1, "<th>" & mains(i) & "</th>")
        Next
        Microsoft.VisualBasic.Print(1, "</tr>")
        '----------------------------------------------------------------------
        For i = 1 To scount 'loop through 2nd layer rows
            Microsoft.VisualBasic.Print(1, "<tr><th>" & subs(i) & "</th>")
            For j = 1 To mcount
                found = False
                first = True
                Microsoft.VisualBasic.Print(1, "<td valign=top>")
                For k = 1 To ecount
                    If subtext(k) = i And maintext(k) = j Then
                        found = True                                  'found=true if there are entries for box
                        If Not first Then                           'first=true for 1st entry as it doesn't need separator
                            Microsoft.VisualBasic.Print(1, separator)
                        Else
                            first = False
                        End If
                        addr = ""
                        use_this_link = False
                        guid = ""
                        If include_link And entries(k).HasHyperlink Then
                            If entries(k).Hyperlink.IsValid Then
                                use_this_link = True
                                If InStr(entries(k).Hyperlink.Address, ":\") > 0 Or InStr(entries(k).Hyperlink.Address, "\\") > 0 Or InStr(entries(k).Hyperlink.Address, "mj-map:/") > 0 Then
                                    addr = entries(k).Hyperlink.Address
                                Else
                                    addr = entries(k).Document.Path & "\" & entries(k).Hyperlink.Address
                                End If
                                guid = entries(k).Hyperlink.TopicBookmarkGuid
                                If addr = "" Then 'blank hyperlink indicates internal link to same map
                                    addr = doc.FullName
                                End If
                            End If
                        End If

                        If shorten_entries And Len(entries(k).Text) > Math.Round(max_length / (mcount + 1)) - 2 Then
                            txt = Left(entries(k).Text, CInt(Math.Round(max_length / (mcount + 1))) - 2) & "..."
                        Else
                            txt = entries(k).Text
                        End If

                        If Len(entries(k).Text) > Math.Round(max_length / (mcount + 1) - 2) Then
                            tasks = entries(k).Text & Chr(13) & Chr(10)
                        Else
                            tasks = ""
                        End If

                        'show 4 layer items in hover text
                        For Each tasktopic In entries(k).AllSubTopics
                            If tasks = "" Then
                                tasks = tasktopic.Text
                            Else
                                tasks = tasks & Chr(13) & Chr(10) & tasktopic.Text
                            End If
                        Next

                        imark = ""
                        omark = ""
                        If mark_no_children And tasks = "" Then
                            imark = "<em>"
                            omark = "</em>"
                        End If
                        If include_notes Then
                            imark = "<h2>" & imark
                            omark = "</h2>" & omark
                        End If
                        If bullet Then imark = "<li>" & imark
                        If include_link And use_this_link Then
                            If Not guid = "" Then
                                prefix = "mj-map:///"
                            Else
                                prefix = ""
                            End If
                            If Not guid = "" Then
                                postfix = "#oid=" & guid2oid(guid)
                            Else
                                postfix = ""
                            End If
                            Microsoft.VisualBasic.Print(1, imark & "<a href=" & Chr(34) & prefix & addr & postfix & Chr(34) & "title=" & Chr(34) & tasks & Chr(34) & ">" & txt & "</a>" & omark)
                        Else
                            Microsoft.VisualBasic.Print(1, imark & txt & omark)
                        End If
                        If include_notes And Not entries(k).Notes.Text = "" Then
                            Microsoft.VisualBasic.Print(1, entries(k).Notes.Text)
                        End If
                    End If
                Next 'k entry
                If Not found Then Microsoft.VisualBasic.Print(1, "&nbsp")
                Microsoft.VisualBasic.Print(1, "</td>")
            Next 'column
            Microsoft.VisualBasic.Print(1, "</tr>")
        Next ' row

        '-----Close out the Table-------------------------------------------------
        Microsoft.VisualBasic.Print(1, "</small></small></table></body></html>")
        Microsoft.VisualBasic.FileClose(1)
        '-----------View the report------------------------------------------------
        If autoopen Then
            Shell("C:\Program Files\Internet Explorer\iexplore.exe " & f, Microsoft.VisualBasic.AppWinStyle.NormalFocus)
        End If
    End Sub


End Module
