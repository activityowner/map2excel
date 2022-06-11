Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Module ao_common
    '12Mar2011 ao_common.mmba generic functions used by ao-tools (e.g. mindreader and mark task complete)
    'Code protected under http://creativecommons.org/licenses/by-nc-nd/3.0/
    'Contact info@activityowner.com for persmission to waive restrictions
    'http://wiki.activityowner.com
    Public mtime As Double
    Public dtime As Double
    Sub VersionCheck(ByRef VersionpageLink As String, ByRef programname As String, ByVal programversion As String, ByRef ConfigDoc As mm.Document)
        'used by various main programs.  Also checks version of ao-common
        Dim VersionCheckFrequency As Double
        If 1 = 0 Then
            VersionCheckFrequency = Val(getoption("versioncheckfrequency", ConfigDoc, Nothing))
            If VersionCheckFrequency > 0 Then
                If DateAdd(DateInterval.Day, VersionCheckFrequency, DateValue(getoption("lastversioncheck", ConfigDoc, Nothing))) <= Today Then
                    If MsgBox("Would you like to check to see if this macro is up to date?", vbYesNo) = vbYes Then
                        Dim ie As Object
                        ie = CreateObject("InternetExplorer.Application")
                        ie.Visible = True
                        ie.navigate(VersionpageLink & "?name=" & programname & "&installed=" & programversion)
                        ie = Nothing
                    End If
                    setoption("lastversioncheck", Str(Today), ConfigDoc)
                    ConfigDoc.Save()
                End If
            End If
        End If
    End Sub
    '--------------------------------------------------------------
    'True/False Checks on Maps and topics
    '--------------------------------------------------------------
    Function f_IsADashboardMap(ByVal m_Doc As mm.Document) As Boolean
        ' Check if a map is a dashboard map
        Const T_uriGRM = "http://schemas.gyronix.com/resultmanager"
        Const T_DashSource = "DashSource" ' source map used to generate dashboard
        Dim s_1 As String
        s_1 = m_Doc.CentralTopic.Attributes(T_uriGRM).GetAttributeValue(T_DashSource) 'Read source: this Is Not Empty If a destination (generated) map
        f_IsADashboardMap = (Len(s_1) > 0) ' has a source path, so is a real dashboard
    End Function
    Function isvisible(ByVal m_app As mm.Application, ByVal tmapname As String) As Boolean
        Dim doc As mm.Document
        isvisible = False
        For Each doc In m_app.VisibleDocuments
            If doc.FullName = tmapname Then
                isvisible = True
                Exit Function
            End If
        Next
        doc = Nothing
    End Function
    Function isopen(ByVal m_app As mm.Application, ByVal tmapname As String) As Boolean
        Dim doc As mm.Document
        isopen = False
        For Each doc In m_app.Documents
            If doc.FullName = tmapname Then
                isopen = True
                Exit Function
            End If
        Next
        doc = Nothing
    End Function
    Function isclone(ByRef tt As mm.Topic, ByRef t As mm.Topic) As Boolean
        isclone = False
        If Not tt.HasHyperlink Then
            isclone = False
        Else
            If (InStr(tt.Hyperlink.Address, "http://mjc.mindjet.com/openlink") > 0) Then
                If (t.Hyperlink.Address = tt.Hyperlink.Address) Then
                    isclone = True
                End If
            Else
                If t.Hyperlink.TopicBookmarkGuid = tt.Hyperlink.TopicBookmarkGuid And Len(tt.Hyperlink.TopicBookmarkGuid) > 0 Then
                    isclone = True
                End If
            End If
        End If
    End Function
    Function isred(ByRef t As mm.Topic) As Boolean
        'used by naa
        Dim red As Byte
        Dim green As Byte
        Dim blue As Byte
        Dim alpha As Byte
        t.TextColor.GetARGB(alpha, red, green, blue)
        isred = (red = 255)
    End Function
    Function isrepeating(ByRef t As mm.Topic) As Boolean
        'used by naa
        Dim i As Integer
        Dim cat As String
        cat = LCase(t.Task.Categories)
        If cat = "" And t.TextLabels.IsValid Then
            For i = 1 To t.TextLabels.Count
                cat = cat & t.TextLabels.Item(i).Name
            Next
        End If
        If cat = "" And Not t.TextLabels.IsValid Then   'this is a hack that will yield false positives
            cat = cat & t.Xml
        End If
        'code below will need modification if additional repeating categories are added
        isrepeating = InStr(cat, "ly") > 0 Or InStr(cat, "day") > 0 Or InStr(cat, "every") > 0 Or InStr(cat, "each") > 0 Or InStr(cat, "endof") > 0 Or InStr(cat, "biannual") > 0
    End Function
    Function parentcontains(ByRef t As mm.Topic, ByVal sometext As String) As Boolean
        'used by naa
        If Not (t.ParentTopic Is Nothing) Then
            parentcontains = InStr(LCase(t.ParentTopic.Text), LCase(sometext)) > 0
        Else
            parentcontains = False
        End If
    End Function

    '--------------------------------------------------------------
    'Sounds
    '--------------------------------------------------------------
    Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

    Sub PlaySoundchirp()
        'Substitute the path and filename of the sound you want to play
        Call sndPlaySound32("c:\windows\media\recycle.wav", 0)
    End Sub
    Function followedhyperlink(ByRef m_app As mm.Application, ByRef t As mm.Topic, ByRef doccurrent As mm.Document) As mm.Topic
        Dim timeout As Double
        Dim d As mm.Document
        Dim tt As mm.Topic
        On Error Resume Next
        If (Not InStr(t.Hyperlink.Address, "http://") > 0) And Len(t.Hyperlink.TopicBookmarkGuid) > 0 Then
            If isabsolute(t.Hyperlink.Address) Then
                d = FindMap(m_app, t.Hyperlink.Address)
                If d Is Nothing Then d = m_app.Documents.Open(t.Hyperlink.Address, "", isvisible(m_app, t.Hyperlink.Address))
            Else
                d = FindMap(m_app, t.Document.Path & t.Hyperlink.Address)
                If d Is Nothing Then d = m_app.Documents.Open(t.Document.Path & t.Hyperlink.Address, "", isvisible(m_app, t.Document.Path & t.Hyperlink.Address))
            End If
            If Not d Is Nothing Then
                For Each tt In d.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
                    If t.Hyperlink.TopicBookmarkGuid = tt.Guid Then
                        followedhyperlink = tt
                        d = Nothing
                        tt = Nothing
                        Exit Function
                    End If
                Next
            End If
        End If
        t.Hyperlink.Follow()
        timeout = 0
        While doccurrent.FullName = m_app.ActiveDocument.FullName
            Pause(0.2)
            timeout = timeout + 1
            If timeout > 50 Then
                If InStr(t.Hyperlink.Address, "http://mjc.mindjet.com/openlink") > 0 Then MsgBox("mjc timeout")
                Exit While
            End If
        End While
        followedhyperlink = m_app.ActiveDocument.Selection.PrimaryTopic
    End Function
    Public Sub Pause(ByVal duration As Double)
        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(duration)
        Do Until Now > dblWaitTil
            Application.DoEvents() ' Allow windows messages to be processed
        Loop
    End Sub
    Function st_create(ByRef tmap As mm.Document, ByRef t As mm.Topic, ByVal stext As String) As mm.Topic
        'find or create a year/month/date branch in completed log
        Dim found As Boolean
        Dim i As Integer
        found = False
        stext = Trim(stext)
        i = t.AllSubTopics.Count 'search from top
        st_create = Nothing
        While Not found And i > 0
            If t.AllSubTopics(i).Text = stext Then
                st_create = t.AllSubTopics(i)
                found = True
                If Val(t.AllSubTopics(i).Text) < Val(stext) Then
                    Exit While
                End If
            End If
            i = i - 1
        End While
        If Not found Then st_create = t.AddSubTopic(stext)
    End Function
    Sub copytocalendarlog(ByRef tmap As mm.Document, ByRef tt As mm.Topic)
        Dim t As mm.Topic
        Dim y As String
        Dim m As String
        Dim d As String
        y = Str(Microsoft.VisualBasic.Year(Today))
        m = Str(Microsoft.VisualBasic.Month(Today))
        d = Mid(Today.ToString, 1, InStr(Today.ToString, " ") - 1)
        t = st_create(tmap, tmap.CentralTopic, "Complete")
        t = st_create(tmap, t, y)
        t = st_create(tmap, t, m)
        t = st_create(tmap, t, d)
        t = t.AddSubTopic("")
        t.Xml = tt.Xml
        t.Task.Complete = 100
        t = Nothing
    End Sub
    Sub copytoRefcalendarlog(ByRef tmap As mm.Document, ByRef tt As mm.Topic, ByRef reftext As String)
        'copy completed task to reference branch of tmap in calendar outline
        Dim t As mm.Topic
        Dim d As String
        t = st_create(tmap, tmap.CentralTopic, reftext)
        If Not t.Icons.HasStockIcon(Mindjet.MindManager.Interop.MmStockIcon.mmStockIconNoEntry) Then t.Icons.AddStockIcon(Mindjet.MindManager.Interop.MmStockIcon.mmStockIconNoEntry)
        t = st_create(tmap, t, Str(Year(Today)))
        t = st_create(tmap, t, Str(Month(Today)))
        d = Mid(Today.ToString, 1, InStr(Today.ToString, " ") - 1)
        t = st_create(tmap, t, d)
        t = t.AddSubTopic("")
        t.Xml = tt.Xml
        t.Task.Complete = 100
        t = Nothing
    End Sub
    '-----------------------------------------------------------------------
    'Configuration Map Functions
    '-----------------------------------------------------------------------------------------
    Sub createoption(ByVal setting As String, ByVal settingvalue As String, ByRef ConfigDoc As mm.Document)
        'create and set option if it isn't already set
        Dim s As mm.Topic
        s = createmainbranch("options", ConfigDoc, "") 'create option branch if necessary
        createkeyword(s, setting, settingvalue, "1", "0")
        s = Nothing
    End Sub
    Sub deleteoption(ByVal setting As String, ByRef ConfigDoc As mm.Document)
        Dim s As mm.Topic
        s = createmainbranch("options", ConfigDoc, "") 'create option branch if necessary
        deletekeyword(s, setting, "1", "0")
        s = Nothing
    End Sub
    Sub setoption(ByVal setting As String, ByVal settingvalue As String, ByRef ConfigDoc As mm.Document)
        'Set option "setting" to "settingvalue" in ConfigDoc
        Dim found As Boolean
        Dim t As mm.Topic
        Dim s As mm.Topic
        Dim ss As mm.Topic
        found = False
        s = createmainbranch("options", ConfigDoc, "")
        found = False
        For Each t In s.AllSubTopics
            If t.Text = setting Then
                found = True
                t.Notes.Text = settingvalue
            End If
        Next
        If Not found Then
            ss = s.AddSubTopic(setting)
            ss.Notes.Text = settingvalue
        End If
        t = Nothing
        s = Nothing
    End Sub
    Function optiontrue(ByRef moption As String, ByRef ConfigDoc As mm.Document, ByRef optionbranch As mm.Topic) As Boolean
        Dim o As String
        o = getoption(moption, ConfigDoc, optionbranch)
        If Trim(o) = "1" Then
            optiontrue = True
        ElseIf Trim(o) = "0" Then
            optiontrue = False
        Else
            If MsgBox("Option:" & moption & " has not been set correctly.  Would you like to set it true?", vbYesNo) = vbYes Then
                createoption(moption, "1", ConfigDoc)
                optiontrue = True
            Else
                createoption(moption, "0", ConfigDoc)
                optiontrue = False
            End If
        End If
    End Function

    Function getoption(ByRef mroption As String, ByRef ConfigDoc As mm.Document, ByRef s As mm.Topic) As String
        'get value for mroption from ConfigDoc
        Dim t As mm.Topic
        If s Is Nothing Then s = createmainbranch("options", ConfigDoc, "")
        getoption = ""
        For Each t In s.AllSubTopics
            If t.Text = mroption Then
                getoption = t.Notes.Text
                Exit Function
            End If
        Next
        Debug.Print("Option " & mroption & " not found in configuration map")
        t = Nothing
        s = Nothing
    End Function
    Sub usersetoption(ByRef ConfigDoc As mm.Document, ByVal ParseText As String)
        'Allow user to set option values with "m"
        Dim opt As String
        Dim optvalue As String
        Dim firstcolon As Integer
        Dim secondcolon As Integer
        firstcolon = InStr(ParseText, ":")
        secondcolon = InStrRev(ParseText, ":")
        opt = Mid(ParseText, firstcolon + 1, secondcolon - firstcolon - 1)
        optvalue = Mid(ParseText, secondcolon + 1)
        setoption(opt, optvalue, ConfigDoc)
        MsgBox("Option " & opt & " set to " & optvalue)
    End Sub
    Sub usergetoption(ByRef ConfigDoc As mm.Document, ByVal ParseText As String)
        'Allow user to get current option values with "m"
        Dim opt As String
        Dim optvalue As String
        Dim firstcolon As Integer
        firstcolon = InStr(ParseText, ":")
        opt = Mid(ParseText, firstcolon + 1)
        optvalue = getoption(opt, ConfigDoc, Nothing)
        MsgBox(opt & " value is " & optvalue)
    End Sub

    Function middlestr(ByVal sometext As String, ByVal starting As Integer, ByVal ending As Integer) As String
        middlestr = Mid(sometext, starting, ending - starting)
    End Function
    Function createmainbranch(ByVal mainstring As String, ByVal ConfigDoc As mm.Document, ByVal callouttext As String) As mm.Topic
        'find or create a main topic for a map

        Dim c As mm.Topic
        createmainbranch = Nothing
        If Not ConfigDoc Is Nothing Then
            createmainbranch = findmainbranch(mainstring, ConfigDoc)
            If createmainbranch Is Nothing Then
                createmainbranch = ConfigDoc.CentralTopic.AddBalancedSubTopic(mainstring)
                If Not callouttext = "" Then
                    c = createmainbranch.CalloutTopics.Add
                    c.Text = callouttext
                    c = Nothing
                End If
            End If
        End If
    End Function
    Function findmainbranch(ByVal mainstring As String, ByVal ConfigDoc As mm.Document) As mm.Topic
        'find main topic for a map if it exists
        Dim i As Integer
        findmainbranch = Nothing
        If Not ConfigDoc Is Nothing Then
            i = ConfigDoc.CentralTopic.AllSubTopics.Count
            While i > 0
                If LCase(RTrim(ConfigDoc.CentralTopic.AllSubTopics(i).Text)) = LCase(RTrim(mainstring)) Then
                    findmainbranch = ConfigDoc.CentralTopic.AllSubTopics(i)
                    Exit Function
                End If
                i = i - 1
            End While
        End If
    End Function
    Sub deletemainbranch(ByVal mainstring As String, ByVal ConfigDoc As mm.Document)
        'find or create a main topic for a map
        Dim found As Boolean
        Dim i As Integer
        i = ConfigDoc.CentralTopic.AllSubTopics.Count
        While i > 0 And Not found
            If LCase(ConfigDoc.CentralTopic.AllSubTopics(i).Text) = LCase(mainstring) Then
                found = True
                ConfigDoc.CentralTopic.AllSubTopics(i).Delete()
            End If
            i = i - 1
        End While
    End Sub

    Sub FixBadOutlinkerLinks(ByRef doc As mm.Document)
        'versions of mindreader before 22Jan08 added outlinker outlook links with "|message" on end if "link" keyword in task
        Dim t As mm.Topic
        For Each t In doc.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
            If t.HasHyperlink Then
                If InStr(t.Hyperlink.Address, "|") > 0 Then
                    t.Hyperlink.Address = Left(t.Hyperlink.Address, InStr(t.Hyperlink.Address, "|") - 1)
                End If
            End If
        Next
        t = Nothing
    End Sub
    Sub createkeyword(ByRef a As mm.Topic, ByRef keyword As String, ByRef code As String, ByVal partofupgrade As String, ByVal lastupgrade As String)
        'creates a new keyword in main branch "a" with value "code".  If it already exists, don't change its value
        Dim b As mm.Topic
        Dim found As Boolean
        If Val(partofupgrade) > Val(lastupgrade) Then
            found = False
            For Each b In a.AllSubTopics
                If LCase(b.Text) = LCase(keyword) Then found = True
            Next
            If Not found Then
                b = a.AddSubTopic(keyword)
                b.Notes.Text = code
            End If
            b = Nothing
        End If
    End Sub
    Sub deletekeyword(ByRef a As mm.Topic, ByRef keyword As String, ByVal partofupgrade As String, ByVal lastupgrade As String)
        'Delete a keyword in main branch a
        Dim b As mm.Topic
        If Val(partofupgrade) > Val(lastupgrade) Then
            For Each b In a.AllSubTopics
                If LCase(b.Text) = LCase(keyword) Then b.Delete()
            Next
            b = Nothing
        End If
    End Sub
    Sub WarnFirstDeleteKeyword(ByRef a As mm.Topic, ByRef keyword As String, ByVal partofupgrade As String, ByVal lastupgrade As String)
        'Delete a keyword in main branch a
        Dim b As mm.Topic
        If Val(partofupgrade) > Val(lastupgrade) Then
            For Each b In a.AllSubTopics
                If LCase(b.Text) = LCase(keyword) Then
                    MsgBox("Note: Program is removing the _" & keyword & "_ keyword from the " & a.Text & " branch to avoid issues")
                    b.Delete()
                End If
            Next
            b = Nothing
        End If
    End Sub
    Sub addkeyword(ByRef a As mm.Topic, ByRef keyword As String, ByRef code As String, ByVal partofupgrade As String, ByVal lastupgrade As String)
        'Add a keyword underneath topic a.  Set value if it exists
        Dim b As mm.Topic
        Dim found As Boolean
        If lastupgrade = "" Then lastupgrade = "0"
        If Val(partofupgrade) > Val(lastupgrade) Then
            found = False
            For Each b In a.AllSubTopics
                If b.Text = keyword Then
                    found = True
                    b.Notes.Text = code
                End If
            Next
            If Not found Then
                b = a.AddSubTopic(keyword)
                b.Notes.Text = code
            End If
            b = Nothing
        End If
    End Sub
    Sub addtriplet(ByRef a As mm.Topic, ByRef keyword As String, ByRef code1 As String, ByRef code2 As String, ByRef code3 As String, ByVal partofupgrade As String, ByVal lastupgrade As String)
        'Add a triplet underneath topic a.  Set value if it exists
        Dim b As mm.Topic
        Dim found As Boolean
        If Val(partofupgrade) > Val(lastupgrade) Then
            found = False
            For Each b In a.AllSubTopics
                If b.Text = keyword Then
                    found = True
                    b.AllSubTopics.Item(1).Delete()
                    b.AllSubTopics.Item(1).Delete()
                    b.AllSubTopics.Item(1).Delete()
                    b.AddSubTopic(code1)
                    b.AddSubTopic(code2)
                    b.AddSubTopic(code3)
                End If
            Next
            If Not found Then
                b = a.AddSubTopic(keyword)
                b.AddSubTopic(code1)
                b.AddSubTopic(code2)
                b.AddSubTopic(code3)
            End If
            b = Nothing
        End If
    End Sub

    Function getmap(ByRef m_app As mm.Application, ByRef mapname As String) As mm.Document
        'opens map named mapname in my maps directory.  Create it if not found
        Dim fullname As String
        Dim createit As Boolean
        If InStr(mapname, ":\") > 0 Or InStr(mapname, "\\") > 0 Then
            fullname = mapname
        Else
            fullname = m_app.GetPath(Mindjet.MindManager.Interop.MmDirectory.mmDirectoryMyMaps) & mapname
        End If
        On Error Resume Next
        getmap = m_app.Documents.Open(fullname, "", isvisible(m_app, fullname))
        On Error GoTo 0
        If getmap Is Nothing Then
            createit = False
            createit = InStr(LCase(fullname), "completed") > 0 And Not InStr(LCase(fullname), "completedconfig") > 0
            If Not createit Then
                createit = MsgBox(fullname & " was not found. Would you like to create it? If this is the first time you are running program or just upgraded, click OK.", vbOKCancel) = vbOK
            End If
            If createit Then
                getmap = m_app.Documents.Add(False)
                On Error Resume Next
                getmap.SaveAs(fullname)
                If Err.Number > 0 Then
                    MsgBox("Error:" & Err.Description)
                    Exit Function
                End If
                On Error GoTo 0
            Else
                Exit Function
            End If
        End If

    End Function
    Sub checkforduplicates(ByVal ConfigDoc As mm.Document)
        'duplicate entries in configuration maps can cause problems.  Check for them upon each upgrade
        Dim m As mm.Topic
        Dim i As Integer
        Dim j As Integer
        Dim notfinished As Boolean
        Dim deletethis As mm.Topic
        notfinished = True
        deletethis = Nothing
        While notfinished
            notfinished = False
            For Each m In ConfigDoc.CentralTopic.AllSubTopics
                If m.AllSubTopics.Count > 1 Then
                    For i = 1 To m.AllSubTopics.Count
                        For j = i To m.AllSubTopics.Count
                            If Not (i = j) And m.AllSubTopics.Item(i).Text = m.AllSubTopics.Item(j).Text Then
                                MsgBox("You had duplicate entries for[" & m.AllSubTopics.Item(i).Text & "] in configuration map branch [" & m.Text & "] at position " & i & " And " & j & ". The 2nd was deleted.")
                                notfinished = True
                                deletethis = m.AllSubTopics.Item(j)
                            End If
                        Next
                    Next
                End If
            Next
            If notfinished Then
                deletethis.Delete()
            End If
        End While
        m = Nothing
        deletethis = Nothing
    End Sub
    Sub copybranchtomap(ByRef Parent As mm.Topic, ByVal Title As String, ByRef destmap As mm.Document)
        'used by naa
        Dim t As mm.Topic
        Dim tt As mm.Topic
        For Each t In Parent.AllSubTopics
            If LCase(t.Text) = LCase(Title) Then
                For Each tt In t.AllSubTopics
                    destmap.CentralTopic.AddSubTopic("").Xml = tt.Xml
                Next
            End If
        Next
        t = Nothing
        tt = Nothing
    End Sub
    Sub copybranchcontainingtomap(ByRef Parent As mm.Topic, ByVal Title As String, ByRef destmap As mm.Document)
        'used by naa
        Dim t As mm.Topic
        Dim tt As mm.Topic
        For Each t In Parent.AllSubTopics
            If InStr(LCase(t.Text), LCase(Title)) > 0 Then
                For Each tt In t.AllSubTopics
                    destmap.CentralTopic.AddSubTopic("").Xml = tt.Xml
                Next
            End If
        Next
        t = Nothing
        tt = Nothing
    End Sub
    Function TotalActivities(ByVal map As mm.Document) As Integer
        'used by naa
        Dim t As mm.Topic
        Dim n As Integer
        n = 0
        For Each t In map.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
            If Not t.Task.IsEmpty Then
                If Not t.IsCalloutTopic Then
                    If t.Task.Complete < 100 Then
                        n = n + 1
                    End If
                End If
            End If
        Next
        TotalActivities = n
    End Function


    Function TotalRedActivities(ByVal map As mm.Document) As Integer
        'used by naa
        Dim t As mm.Topic
        Dim n As Integer
        n = 0
        For Each t In map.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
            If Not t.Task.IsEmpty Then
                If Not t.IsCalloutTopic Then
                    If isred(t) Then
                        If t.Task.Complete < 100 Then
                            n = n + 1
                        End If
                    End If
                End If
            End If
        Next
        TotalRedActivities = n
        t = Nothing
    End Function
    Function TotalRedCallouts(ByVal map As mm.Document) As Integer
        'used by naa
        Dim t As mm.Topic
        Dim n As Integer
        n = 0
        For Each t In map.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
            If Not t.Task.IsEmpty Then
                If t.IsCalloutTopic Then
                    If isred(t) Then
                        If t.Task.Complete < 100 Then
                            n = n + 1
                        End If
                    End If
                End If
            End If
        Next
        TotalRedCallouts = n
        t = Nothing
    End Function
    Function TotalRedActivitiesWithParentContaining(ByVal map As mm.Document, ByVal s As String) As Integer
        'used by naa
        Dim t As mm.Topic
        Dim n As Integer
        n = 0
        For Each t In map.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
            If Not t.Task.IsEmpty Then
                If Not t.IsCalloutTopic Then
                    If InStr(t.ParentTopic.Text, s) > 0 Then
                        If isred(t) Then
                            n = n + 1
                        End If
                    End If
                End If
            End If
        Next
        TotalRedActivitiesWithParentContaining = n
        t = Nothing
    End Function
    Function TotalActivitiesWithParentContainingandnoduedate(ByVal map As mm.Document, ByVal s As String) As Integer
        'used by naa
        Dim t As mm.Topic
        Dim n As Integer
        n = 0
        For Each t In map.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
            If Not t.Task.IsEmpty Then
                If Not t.IsCalloutTopic Then
                    If InStr(t.ParentTopic.Text, s) > 0 Then
                        If isdate0(t.Task.DueDate) Then
                            If t.CalloutTopics.Count > 0 Then
                                If isdate0(t.CalloutTopics.Item(1).Task.DueDate) Then
                                    n = n + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        TotalActivitiesWithParentContainingandnoduedate = n
        t = Nothing
    End Function
    Function TotalActivitiesWithParentContaining(ByVal map As mm.Document, ByVal s As String) As Integer
        'used by naa
        Dim t As mm.Topic
        Dim n As Integer
        n = 0
        For Each t In map.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
            If Not t.Task.IsEmpty Then
                If Not t.IsCalloutTopic Then
                    If InStr(t.ParentTopic.Text, s) > 0 Then
                        n = n + 1
                    End If
                End If
            End If
        Next
        TotalActivitiesWithParentContaining = n
        t = Nothing
    End Function
    Function arrayaverage(ByVal age As Object, ByVal numdatedactions As Integer) As Double
        'used by naa
        Dim i As Integer
        arrayaverage = 0
        For i = 1 To numdatedactions
            arrayaverage = arrayaverage + CDbl(age(i)) / numdatedactions
        Next
    End Function
    Function maxbranch(ByRef doc As mm.Document) As mm.Topic
        'used by naa
        Dim t As mm.Topic
        Dim st As mm.Topic
        Dim root As Boolean
        Dim maxcount As Integer
        maxcount = 0
        maxbranch = Nothing
        For Each t In doc.CentralTopic.AllSubTopics
            If t.AllSubTopics.Count > 0 Then
                root = True
                For Each st In t.AllSubTopics
                    If st.AllSubTopics.Count > 0 Then
                        root = False
                        If st.AllSubTopics.Count > maxcount Then
                            maxbranch = st
                            maxcount = t.AllSubTopics.Count
                        End If
                    End If
                Next
                If root = True Then
                    If t.AllSubTopics.Count > maxcount Then
                        maxbranch = t
                        maxcount = t.AllSubTopics.Count
                    End If
                End If
            End If
        Next
        t = Nothing
        st = Nothing
    End Function
    Function cat(ByRef t As mm.Topic) As String
        'used by naa
        Dim i As Integer
        cat = LCase(t.Task.Categories)
        If cat = "" And t.TextLabels.IsValid Then
            For i = 1 To t.TextLabels.Count
                cat = cat & ", " & t.TextLabels.Item(i).Name

            Next
        End If
        If gethidcat(t) = "" Then
            If cat = "" And Not t.TextLabels.IsValid Then   'this is a hack that will yield false positives
                cat = cat & ", " & t.Xml
            End If
        Else
            cat = gethidcat(t) & ", " & cat
        End If
    End Function
    Function gethidcat(ByVal t As mm.Topic) As String
        'used by naa
        Dim s As String
        s = ""
        If InStr(t.Xml, "HidCats") > 0 Then
            s = Mid(t.Xml, InStr(t.Xml, "HidCats=") + 9)
            s = Mid(s, 1, InStr(s, Chr(34)) - 1)
        End If
        If Len(s) = 0 And InStr(t.Xml, "MirCat") > 0 Then
            s = Mid(t.Xml, InStr(t.Xml, "MirCat=") + 8)
            s = Mid(s, 1, InStr(s, Chr(34)) - 1)
        End If
        gethidcat = s
    End Function
    Function getfirstarea(ByRef t As mm.Topic) As String
        'used by naa
        Dim cats As String
        Dim start As Integer
        Dim comma As Integer
        cats = cat(t)
        getfirstarea = ""
        If InStr(cats, "^") > 0 Then
            cats = Mid(cats, InStr(cats, "^") + 1)
            If InStr(cats, ";") > 0 Then
                getfirstarea = Mid(cats, 1, InStr(cats, ";") - 1)
            ElseIf InStr(cats, ",") > 0 Then
                getfirstarea = Mid(cats, 1, InStr(cats, ",") - 1)
            Else
                getfirstarea = cats
            End If
        Else
            cats = gethidcat(t)
            start = InStr(cats, "^")
            If start > 0 Then
                comma = InStr(Mid(cats, start), ";")
                If Not comma > 0 Then
                    comma = InStr(Mid(cats, start), ",")
                End If
            Else
                comma = 0
            End If
            If start > 0 Then
                If comma > 0 Then
                    getfirstarea = Mid(cats, start + 1, comma - 2)
                Else
                    getfirstarea = Mid(cats, start + 1)
                End If
            End If
        End If
    End Function
    Function GetNewVersion(ByRef m_app As mm.Application, ByRef ProgramVersion As String, ByRef VersionMapLink As String) As Boolean
        'generic version checker
        Dim VersionDoc As mm.Document
        Dim releasedversion As String
        GetNewVersion = False
        On Error GoTo x1
        VersionDoc = m_app.Documents.Open(VersionMapLink, "", False)
        On Error GoTo x2
        releasedversion = VersionDoc.CentralTopic.Notes.Text
        On Error GoTo 0
        If Err.Number > 0 Then
x1:         Debug.Print("Version map not accessible") : Exit Function
x2:         Debug.Print("Could not get current version number from " & VersionMapLink) : Exit Function
        End If
        'Debug.Print(Val(ProgramVersion))
        'Debug.Print(Val(releasedversion))
        'Debug.Print(Val(ProgramVersion) - Val(releasedversion))
        If Val(ProgramVersion) < Val(releasedversion) Then
            GetNewVersion = MsgBox("You currently have version " & ProgramVersion & " and the most recent version is " & releasedversion & ".  Would you like to visit the upgrade page?", vbYesNo) = vbYes
        Else
            GetNewVersion = False
            Debug.Print("You are using version " & ProgramVersion & " and the latest released version is " & releasedversion & ". You appear To be up to date.")
        End If
        If GetNewVersion Then VersionDoc.CentralTopic.Hyperlink.Follow()
        VersionDoc.Close()
    End Function
    Function issaved(ByRef doc As mm.Document) As Boolean
        'used by mindreaderopen
        issaved = isworkspacemap(doc) Or Not doc.FullName = doc.Name
    End Function
    Function isworkspacemap(ByRef doc As mm.Document) As Boolean
        'used by mindreaderopen
        Dim prefix As String
        Dim issaved As Boolean
        'attempt to determine if has been saved or is a workspace map
        prefix = "Map"
        issaved = True
        If Left(doc.FullName, Len(prefix)) = prefix Then
            issaved = Not IsNumeric(Right(doc.FullName, Len(doc.FullName) - Len(prefix)))
        End If
        isworkspacemap = (doc.FullName = doc.Name) And issaved
    End Function
    Function isabsolute(ByRef fname As String) As Boolean
        'determine if a path is relative or absolute
        isabsolute = InStr(fname, "\\") > 0 Or InStr(fname, ":\") > 0
    End Function
    Function ismindjetconnectlink(ByRef link As mm.Hyperlink) As Boolean
        ismindjetconnectlink = InStr(LCase(link.Address), "http://mjc.mindjet.com/openlink") = 1
    End Function
    Sub CloseHiddenMaps(ByRef m_app As mm.Application)
        'historically routines have left maps open hidden to improve performance but
        'this leads to conflicts when using multi-computer setups
        Dim doc As mm.Document
        For Each doc In m_app.Documents(True)
            If Not isvisible(m_app, doc.FullName) Then
                If doc.IsModified Then doc.Save()
                If InStr(LCase(doc.Name), "config") = 0 Then doc.Close()
            End If
        Next
    End Sub

    Function OpenMapHidden(ByRef m_app As mm.Application, ByRef mapname As String) As mm.Document
        'open map hidden unless already visible
        sw("in openmaphidden")
        OpenMapHidden = FindMap(m_app, mapname)
        If OpenMapHidden Is Nothing Then
            OpenMapHidden = m_app.Documents.Open(mapname, "", False)
        End If
        sw("leaving openmaphidden")
    End Function
    Function FindMap(ByRef m_app As mm.Application, ByRef mapname As String) As mm.Document
        Dim doc As mm.Document
        For Each doc In m_app.Documents
            If doc.FullName = mapname Then
                FindMap = doc
                doc = Nothing
                Exit Function
            End If
        Next
        FindMap = Nothing
    End Function
    Function LinkToThisTopic(ByRef t As mm.Topic) As String
        LinkToThisTopic = "mj-map:///" & Replace(t.Document.FullName, " ", "%20") & "#oid=" & guid2oid(t.Guid)
    End Function
    Function LinktoThisTopicHyperlink(ByRef t As mm.Topic) As String
        Dim prefix As String
        Dim postfix As String
        Dim addr As String
        Dim guid As String
        prefix = ""
        postfix = ""
        addr = ""
        guid = ""
        If t.HasHyperlink Then
            If t.Hyperlink.IsValid Then
                If InStr(t.Hyperlink.Address, ":\") > 0 Or InStr(t.Hyperlink.Address, "\\") > 0 Or InStr(t.Hyperlink.Address, "mj-map:/") > 0 Or InStr(t.Hyperlink.Address, "Outlook") > 0 Or InStr(t.Hyperlink.Address, "http") > 0 Then
                    addr = t.Hyperlink.Address
                Else
                    addr = t.Document.Path & "\" & t.Hyperlink.Address
                End If
                guid = t.Hyperlink.TopicBookmarkGuid
                If addr = "" Then 'blank hyperlink indicates internal link to same map
                    addr = t.Document.FullName
                End If
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
            End If
        End If
        If InStr(t.Hyperlink.Address, "https") > 0 Then
            LinktoThisTopicHyperlink = t.Hyperlink.Address
        Else
            LinktoThisTopicHyperlink = prefix & addr & postfix
        End If
    End Function

    Function hyperlinkstring(ByRef t As mm.Topic) As String
        If t.HasHyperlink Then
            hyperlinkstring = "<a href=" & Chr(34) & LinktoThisTopicHyperlink(t) & Chr(34) & ">" & t.Text & "</a>"
        Else
            hyperlinkstring = "<a href=" & Chr(34) & LinkToThisTopic(t) & Chr(34) & ">" & t.Text & "</a>"
        End If
    End Function

    Function guid2oid(ByVal base64String As String) As String
        '28Nov08 http://creativecommons.org/licenses/by-sa/2.5/ http://www.activityowner.com
        'convert topic.guid to oid
        'Derived from: 1999 Antonin Foller, Motobit Software, http://www.motobit.com/tips/detpg_Base64/
        Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        Dim ngroupall As String
        Dim dataLength As Integer
        Dim groupBegin As Integer
        dataLength = Len(base64String)
        ngroupall = ""
        For groupBegin = 1 To dataLength Step 4
            Dim numDataBytes As Integer
            Dim CharCounter As Integer
            Dim thisChar As String
            Dim thisData As Integer
            Dim nGroupInt As Integer
            Dim nGroupStr As String
            numDataBytes = 3
            nGroupInt = 0
            For CharCounter = 0 To 3
                thisChar = Mid(base64String, groupBegin + CharCounter, 1)
                If thisChar = "=" Then
                    numDataBytes = numDataBytes - 1
                    thisData = 0
                Else
                    thisData = InStr(1, Base64, thisChar) - 1
                End If
                nGroupInt = 64 * nGroupInt + thisData
            Next
            nGroupStr = Hex(nGroupInt)
            nGroupStr = mystring(6 - Len(nGroupStr), "0") & nGroupStr
            ngroupall = ngroupall & nGroupStr
        Next
        guid2oid = "{" & _
                Mid(ngroupall, 7, 2) & Mid(ngroupall, 5, 2) & Mid(ngroupall, 3, 2) & Mid(ngroupall, 1, 2) & "-" & _
                Mid(ngroupall, 11, 2) & Mid(ngroupall, 9, 2) & "-" & _
                Mid(ngroupall, 15, 2) & Mid(ngroupall, 13, 2) & "-" & _
                Mid(ngroupall, 17, 2) & Mid(ngroupall, 19, 2) & "-" & _
                Mid(ngroupall, 21, 12) & "}"
    End Function
    Function mystring(ByVal mylen As Integer, ByVal mystr As String) As String
        Dim i As Integer
        mystring = ""
        For i = 1 To mylen
            mystring = mystring & Mid(mystr, 1, 1)
        Next
    End Function
    Sub sw(ByVal label As String)
        'uncomment below for benchmarking
        ' If mtime = 0 Then mtime = Now
        'Debug.Print(Round(Now - mtime, 2) & "   :   " & Round(Timer - dtime, 2) & "      :" & label)
        'dtime = Timer
    End Sub
    Sub setstartdate(ByRef t As mm.Topic, ByVal d As Date)
        'MindManager 9 behaves strangely with regard to setting start dates.  If a due date is present, it will assume a duration
        'of 0 days and move the due date to the start date being set.   If a start and due date is present, it will move the due date back
        'according to the amount the start date is being moved.  This function attempts to work around that behavior by checking the state
        'of the task prior to adjusting the start date and then adjusting accordingly.
        '
        'scenarios:
        '1: no start or due date present
        '2: only start date present
        '3: only due date present
        '4: start and due date present
        'for all these scenarios, the easiest strategy is to save the due date, set the start date, and then fix the due date

        Dim temp As Date
        'If d = 0 And t.Task.StartDate > 0 And t.Task.DueDate > 0 Then
        'MsgBox("Start Date can not be removed if start date and end date are already specified")
        'End If
        If Not t.Task.StartDateReadOnly Then
            temp = t.Task.DueDate
            t.Task.StartDate = d
            t.Task.DueDate = temp
        Else
            MsgBox("Start date is readonly")
        End If
    End Sub
    Function eval(ByRef m_app As Mindjet.MindManager.Interop.Application, ByVal s As String) As String
        Microsoft.VisualBasic.FileOpen(1, "c:\windows\temp\mindreader.tmp", OpenMode.Output, OpenAccess.Write)
        Microsoft.VisualBasic.Print(1, s)
        Microsoft.VisualBasic.FileClose(1)
        m_app.RunMacro(Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com AORibbon\evalclipboard.mmbas")
        eval = Replace(My.Computer.FileSystem.ReadAllText("c:\windows\temp\mindreader.tmp"), Chr(34), "")
    End Function
    Function inteval(ByRef m_app As Mindjet.MindManager.Interop.Application, ByVal s As String) As Integer
        Microsoft.VisualBasic.FileOpen(1, "c:\windows\temp\mindreader.tmp", OpenMode.Output, OpenAccess.Write)
        Microsoft.VisualBasic.Print(1, s)
        Microsoft.VisualBasic.FileClose(1)
        m_app.RunMacro(Environ("ProgramFiles") & "\ActivityOwner.Com\ActivityOwner.Com AORibbon\evalclipboard.mmbas")
        s = My.Computer.FileSystem.ReadAllText("c:\windows\temp\mindreader.tmp")
        s = Replace(s, Chr(34), "")
        inteval = CInt(s)
    End Function
    Function isdate0(ByRef d As Date) As Boolean
        If DateDiff(DateInterval.Day, DateSerial(1899,12,31), d) = -1 Then
            isdate0 = True
        Else
            isdate0 = False
        End If
    End Function
    'Sub donatedelay()
    '   If Not donated Then
    '      Pause(1)
    '     MsgBox("Visit the donate page to eliminate this delay and dialog box")
    'End If
    'End Sub
    Function RMinstalled() As Boolean
        Dim regKey As RegistryKey
        Dim keyValue As String
        keyValue = "SOFTWARE"
        regKey = Registry.CurrentUser.OpenSubKey(keyValue, False)
        If Not regKey Is Nothing Then
            regKey = regKey.OpenSubKey("Gyronix")
        End If
        If Not regKey Is Nothing Then
            regKey = regKey.OpenSubKey("ResultsManager")
        End If
        If regKey Is Nothing Then
            RMinstalled = False
        Else
            RMinstalled = True
        End If
        If Not regKey Is Nothing Then regKey.Close()
    End Function
End Module
