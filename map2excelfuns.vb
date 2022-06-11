Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Module map2excelfuns
    Function getparent(ByRef m_app As mm.Application) As mm.Topic
        getparent = m_app.ActiveDocument.CentralTopic
        If m_app.ActiveDocument.Selection.Count > 0 And Not m_app.ActiveDocument.Selection.PrimaryTopic Is m_app.ActiveDocument.CentralTopic Then
            If MsgBox("Do you want to export just the selected branch?", vbYesNo) = vbYes Then getparent = m_app.ActiveDocument.Selection.PrimaryTopic
        End If
    End Function
    Function openexcelsheet() As excel.Worksheet
        Dim excelapp As excel.Application
        Dim excelworkbook As excel.Workbook
        excelapp = CType(CreateObject("excel.Application"), excel.Application)
        excelworkbook = excelapp.Workbooks.Add
        excelapp.Visible = True
        openexcelsheet = CType(excelworkbook.Sheets.Item(1), excel.Worksheet)
        excelapp = Nothing
        excelworkbook = Nothing
    End Function
    Sub adjustcolwidth(ByRef sheet As excel.Worksheet, ByRef t As mm.Topic, ByVal maxcolnum As Integer, ByVal colnum As Integer, ByRef rng As excel.Range)
        If colnum > maxcolnum Then setcolwidth(sheet, colnum, 80)
        If (Len(t.Text) > rng.ColumnWidth) Then
            If Len(t.Text) < 80 Then
                setcolwidth(sheet, colnum, Len(t.Text))
            Else
                setcolwidth(sheet, colnum, 80)
            End If
        End If
    End Sub
    Sub putnotesincomment(ByRef rng As excel.Range, ByVal notes As String)
        If getmtckey("options", "notesincomments") = "1" And Len(notes) > 0 Then
            Try
                With rng
                    .Comment.Text(Text:=notes)
                    .Comment.Shape.ScaleHeight(5, 0)
                    .Comment.Shape.ScaleWidth(4, 0)
                End With
            Catch
            End Try
        End If
    End Sub

    Sub addhyperlinkentry(ByRef sheet As excel.Worksheet, ByRef t As mm.Topic, ByRef rng As excel.Range, ByVal topictext As String)
        'If (getmtckey("options", "addhyperlinks") = "1") And t.HasHyperlink Then rng.Value = HyperlinkEntry(t, Replace(topictext, Chr(34), ""))
        Dim tempstate As Boolean
        If (getmtckey("options", "addhyperlinks") = "1") And t.HasHyperlink Then
            rng.Value = t.Text
            tempstate = t.Hyperlink.Absolute
            t.Hyperlink.Absolute = True
            'rng.Hyperlinks.Add(rng, t.Hyperlink.Address)
            Debug.Print("adding hyperlink to excel")
            Debug.Print(t.Hyperlink.Address)
            Debug.Print(t.Text)
            Dim formulastring As String
            formulastring = "=hyperlink(" & Chr(34) & t.Hyperlink.Address & Chr(34) & "," & Chr(34) & t.Text & Chr(34) & ")"
            Debug.Print(formulastring)
            rng.Formula = formulastring
            t.Hyperlink.Absolute = tempstate
        End If
    End Sub

    Sub AddTopicLink(ByRef rng As excel.Range, ByVal t As mm.Topic, ByVal caption As String)
        Dim topicentry As String
        caption = Replace(caption, Chr(34), "")
        Try
            If Len(caption) < 254 Then
                topicentry = "=Hyperlink(" & Chr(34) & LinkToThisTopic(t) & Chr(34) & "," & Chr(34) & caption & Chr(34) & ")"
            Else
                topicentry = caption
            End If
            rng.Formula = topicentry
        Catch
            MsgBox("ERROR WITH " & caption)
            topicentry = caption
        End Try
    End Sub

    Sub narrowupcolumns(ByRef sheet As excel.Worksheet, ByVal maxcolnum As Integer, ByVal rownum As Integer)
        Dim maxwidth As Integer
        Dim startrow As Integer
        Dim i As Integer
        Dim j As Integer
        startrow = headerrow + 1

        For i = 1 To maxcolnum
            sheet.Range(convertaddress(headerrow, i)).Orientation = 90
        Next

        For i = 1 To maxcolnum
            maxwidth = 1
            For j = startrow To rownum
                If Len(sheet.Cells(j, i).text) > maxwidth Then
                    maxwidth = Len(sheet.Cells(j, i).text)
                End If
            Next
            If maxwidth > 80 Then maxwidth = 80
            setcolwidth(sheet, i, maxwidth + 1)
        Next

        If hasimages Then setcolwidth(sheet, imagecol, 10)

        If getmtckey("options", "wraptext") = "1" Then
            For i = 1 To maxcolnum
                For j = 2 To rownum
                    sheet.Range(convertaddress(j, i)).WrapText = True
                Next
            Next
        End If
    End Sub

    Sub copyfontcolortoExcel(ByRef t As mm.Topic, ByRef rng As excel.Range)
        Dim palpha As Byte
        Dim pred As Byte
        Dim pgreen As Byte
        Dim pblue As Byte
        t.TextColor.GetARGB(palpha, pred, pgreen, pblue)
        If Not (pred = 255 And pgreen = 255 And pblue = 255) Then 'don't change to white
            rng.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(pred, pblue, pgreen))
        End If
        'didn't work
        't.FillColor.GetARGB(palpha, pred, pgreen, pblue)
        'rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(pred, pblue, pgreen))
    End Sub

    Sub copyfontcolortoMap(ByRef t As mm.Topic, ByRef rng As excel.Range)
        'Dim palpha As Byte
        'Dim pred As Byte
        'Dim pgreen As Byte
        'Dim pblue As Byte
        'Dim color As System.Drawing.Color
        '  Color = System.Drawing.ColorTranslator.FromOle(rng.Font.Color)
        '= System.Drawing.ColorTranslator.ToOle(Color.FromArgb(pred, pblue, pgreen))

        'If Not (pred = 255 And pgreen = 255 And pblue = 255) Then 'don't change to white
        't.TextColor.SetARGB(palpha, pred, pgreen, pblue)

    End Sub

    Function HyperlinkEntry(ByRef t As mm.Topic, ByVal caption As String) As String
        If Len(caption) < 254 Then
            HyperlinkEntry = "=Hyperlink(" & Chr(34) & LinktoThisTopicHyperlink(t) & Chr(34) & "," & Chr(34) & caption & Chr(34) & ")"
        Else
            HyperlinkEntry = caption
        End If

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
            If InStr(t.Hyperlink.Address, "https") > 0 Then
                LinktoThisTopicHyperlink = t.Hyperlink.Address
            Else
                LinktoThisTopicHyperlink = prefix & addr & postfix
            End If
        Else
            LinktoThisTopicHyperlink = ""
        End If

    End Function
    Function mystring(ByVal mylen As Integer, ByVal mystr As String) As String
        Dim i As Integer
        mystring = ""
        For i = 1 To mylen
            mystring = mystring & Mid(mystr, 1, 1)
        Next
    End Function


    Sub InsertPicture(ByVal PictureFileName As String, ByVal TargetCell As excel.Range, _
        ByVal CenterH As Boolean, ByVal CenterV As Boolean, ByRef activesheet As excel.Workbook)
        ' inserts a picture at the top left position of TargetCell
        ' the picture can be centered horizontally and/or vertically
        Dim p As Object, t As Double, l As Double, w As Double, h As Double
        If TypeName(activesheet) <> "Worksheet" Then Exit Sub
        If Dir(PictureFileName) = "" Then Exit Sub
        ' import picture
        p = activesheet.Pictures.Insert(PictureFileName)
        ' determine positions
        With TargetCell
            t = CType(.Top, Double)
            l = CType(.Left, Double)
            If CenterH Then
                w = CType(CType(.Offset(0, 1).Left, Double) - l, Double)
                l = l + w / 2 - CType(p.Width, Double) / 2
                If l < 1 Then l = 1
            End If
            If CenterV Then
                h = CType(.Offset(1, 0).Top, Double) - t
                t = t + h / 2 - CType(p.Height, Double) / 2
                If t < 1 Then t = 1
            End If
        End With
        ' position picture
        With p
            .Top = t
            .Left = l
        End With
        p = Nothing
    End Sub
    Sub setcolwidth(ByRef sheet As excel.Worksheet, ByVal colnum As Integer, ByVal colwidth As Integer)
        sheet.Range(convertaddress(1, colnum)).ColumnWidth = colwidth
    End Sub
    Sub setrowheight(ByRef sheet As excel.Worksheet, ByVal rownum As Integer, ByVal rowheight As Integer)
        sheet.Range(convertaddress(rownum, 1)).RowHeight = rowheight
    End Sub
    Sub InsertPictureInComment(ByVal PictureFileName As String, ByVal TargetCells As excel.Range, ByRef activesheet As Object)
        Dim c As excel.Comment
        c = TargetCells.AddComment("")
        Try
            c.Shape.Fill.UserPicture(PictureFileName)
            c.Shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
            c.Shape.Width = c.Shape.Width * 3

        Catch
            MsgBox("Error putting picture in comment: " & PictureFileName & Err.Description)
        End Try
        c = Nothing
    End Sub
    Sub InsertPictureInRange(ByVal PictureFileName As String, ByVal TargetCells As excel.Range, ByRef activesheet As Object)
        ' inserts a picture and resizes it to fit the TargetCells range
        Dim p As Object, t As Double, l As Double, w As Double, h As Double
        If TypeName(activesheet) <> "Worksheet" Then Exit Sub
        If Dir(PictureFileName) = "" Then Exit Sub
        ' import picture
        hasimages = True
        p = activesheet.Pictures.Insert(PictureFileName)
        ' determine positions
        With TargetCells
            t = CType(.Top, Double)
            l = CType(.Left, Double)
            w = CType(.Offset(0, .Columns.Count).Left, Double) - l
            h = CType(.Offset(.Rows.Count, 0).Top, Double) - t
        End With
        ' position picture
        With p
            .Top = t
            .Left = l
            .Width = w
            .Height = h
        End With
        p = Nothing
    End Sub

    Function opttrue(ByRef optionstr As String) As Boolean
        opttrue = getmtckey("options", optionstr) = "1"
    End Function

    Function custompropertystringlist(ByRef doc As mm.Document) As ArrayList
        Dim t As mm.Topic
        Dim p As mm.CustomProperty
        Dim custompropertystrings As New ArrayList
        For Each t In doc.Range(Mindjet.MindManager.Interop.MmRange.mmRangeAllTopics)
            If t.DataContainer.DataContainerType = mm.MmDataContainerType.mmDataContainerTypeCustomProperties Then
                For Each p In t.DataContainer.CustomProperties.CustomPropertyCollection
                    If Not custompropertystrings.Contains(p.CustomPropertyName) Then
                        custompropertystrings.Add(p.CustomPropertyName)
                    End If
                Next
            End If
        Next
        custompropertystringlist = custompropertystrings
    End Function

    Function hascustomproperty(ByRef t As mm.Topic, ByRef s As String) As Boolean
        Dim p As mm.CustomProperty
        hascustomproperty = False
        If t.DataContainer.DataContainerType = mm.MmDataContainerType.mmDataContainerTypeCustomProperties Then
            For Each p In t.DataContainer.CustomProperties.CustomPropertyCollection
                If p.CustomPropertyName = s Then
                    hascustomproperty = True
                    Exit For
                End If
            Next
        End If
    End Function
    Function getcustompropertyvalue(ByRef t As mm.Topic, ByRef s As String) As String
        Dim p As mm.CustomProperty
        getcustompropertyvalue = ""
        If t.DataContainer.DataContainerType = mm.MmDataContainerType.mmDataContainerTypeCustomProperties Then
            For Each p In t.DataContainer.CustomProperties.CustomPropertyCollection
                If p.CustomPropertyName = s Then
                    getcustompropertyvalue = p.Value.ToString
                    Exit For
                End If
            Next
        End If
    End Function

    Function IconStringList(ByRef doc As mm.Document) As ArrayList 'bug -- limit by filter?
        Dim iconstrings As ArrayList
        Dim t As mm.Topic
        Dim icn As mm.Icon

        iconstrings = New ArrayList
        For Each t In doc.Range(mm.MmRange.mmRangeAllTopics)
            For Each icn In t.AllIcons
                If Not iconstrings.Contains(GetIconGroupName(doc, icn) & ":" & GetIconDisplayName(doc, icn)) Then
                    iconstrings.Add(GetIconGroupName(doc, icn) & ":" & GetIconDisplayName(doc, icn))
                End If
            Next
        Next
        IconStringList = iconstrings
    End Function
    Function tagstringList(ByRef doc As mm.Document) As ArrayList 'bug -- limit by filter?
        Dim tagstrings As ArrayList
        Dim t As mm.Topic
        tagstrings = New ArrayList
        For Each t In doc.Range(mm.MmRange.mmRangeAllTopics)
            For i = 1 To t.TextLabels.Count
                If Not tagstrings.Contains(t.TextLabels.Item(i).Name) Then
                    tagstrings.Add(t.TextLabels.Item(i).Name)
                End If
            Next
        Next
        tagstringList = tagstrings
    End Function


    Function hasicon(ByRef doc As mm.Document, ByRef t As mm.Topic, ByVal s As String) As Boolean 'bug -- limit by filter?
        Dim icn As mm.Icon
        Dim gname As String
        Dim iname As String
        gname = Mid(s, 1, InStr(s, ":") - 1)
        iname = Mid(s, InStr(s, ":") + 1, Len(s))
        hasicon = False
        For Each icn In t.AllIcons
            If GetIconDisplayName(doc, icn) = iname And GetIconGroupName(doc, icn) = gname Then
                hasicon = True
                Exit For
            End If
        Next
    End Function
    Function convertaddress(ByVal r As Integer, ByVal c As Integer) As String
        Dim s As String
        s = "A"
        If c <= 26 Then
            s = "" & Chr(c - 26 * 0 + 64)
        ElseIf c <= 26 * 2 Then
            s = "A" & Chr(c - 26 * 1 + 64)
        ElseIf c <= 26 * 3 Then
            s = "B" & Chr(c - 26 * 2 + 64)
        ElseIf c <= 26 * 4 Then
            s = "C" & Chr(c - 26 * 3 + 64)
        ElseIf c <= 26 * 5 Then
            s = "D" & Chr(c - 26 * 4 + 64)
        ElseIf c <= 26 * 6 Then
            s = "E" & Chr(c - 26 * 5 + 64)
        ElseIf c <= 26 * 7 Then
            s = "F" & Chr(c - 26 * 6 + 64)
        ElseIf c <= 26 * 8 Then
            s = "G" & Chr(c - 26 * 7 + 64)
        End If
        convertaddress = s & Trim(Str(r))
    End Function

    Function GetIconGroupName(ByRef doc As mm.Document, ByRef ic As mm.Icon) As String

        Dim allGroups As mm.MapMarkerGroups
        Dim aGroup As mm.MapMarkerGroup
        Dim aMapMarker As mm.MapMarker
        Dim result As String

        result = ic.Name    ' set the default return
        Try
            If (ic.Type = mm.MmIconType.mmIconTypeCustom) Then
                '  spin through map markers to find a display name using icon signatures
                allGroups = doc.MapMarkerGroups
                ' Debug.Clear
                For Each aGroup In allGroups
                    ' Debug.Print "======"
                    ' Debug.Print "Group " + aGroup.Name + " contains:"
                    For Each aMapMarker In aGroup
                        ' Debug.Print aMapMarker.Label + " == " + aMapMarker.Icon.Name
                        If (aMapMarker.Icon.Type = mm.MmIconType.mmIconTypeCustom) Then ' only mmIconTypeCustom types have CustomIconSignature properties
                            If StrComp(aMapMarker.Icon.CustomIconSignature, ic.CustomIconSignature) = 0 Then
                                result = aGroup.Name   ' found a display name so exit
                                GoTo ExitNow
                            End If
                        End If
                    Next
                Next
            Else
                ' otherwise spin through map markers using icon names to find a display name
                allGroups = doc.MapMarkerGroups
                ' Debug.Clear
                For Each aGroup In allGroups
                    ' Debug.Print "======"
                    ' Debug.Print "Group " + aGroup.Name + " contains:"
                    For Each aMapMarker In aGroup
                        ' Debug.Print aMapMarker.Label + " == " + aMapMarker.Icon.Name
                        If StrComp(aMapMarker.Icon.Name, ic.Name) = 0 Then
                            result = aGroup.Name   ' found a display name so exit
                            GoTo ExitNow
                        End If
                    Next
                Next
            End If

        Catch
        End Try

ExitNow:
        GetIconGroupName = result
    End Function
    Function GetIconDisplayName(ByRef doc As mm.Document, ByRef ic As mm.Icon) As String

        Dim allGroups As mm.MapMarkerGroups
        Dim aGroup As mm.MapMarkerGroup
        Dim aMapMarker As mm.MapMarker
        Dim result As String

        result = ic.Name    ' set the default return
        Try
            If (ic.Type = mm.MmIconType.mmIconTypeCustom) Then
                '  spin through map markers to find a display name using icon signatures
                allGroups = doc.MapMarkerGroups
                ' Debug.Clear
                For Each aGroup In allGroups
                    ' Debug.Print "======"
                    ' Debug.Print "Group " + aGroup.Name + " contains:"
                    For Each aMapMarker In aGroup
                        ' Debug.Print aMapMarker.Label + " == " + aMapMarker.Icon.Name
                        If (aMapMarker.Icon.Type = mm.MmIconType.mmIconTypeCustom) Then ' only mmIconTypeCustom types have CustomIconSignature properties
                            If StrComp(aMapMarker.Icon.CustomIconSignature, ic.CustomIconSignature) = 0 Then
                                result = aMapMarker.Label   ' found a display name so exit
                                GoTo ExitNow
                            End If
                        End If
                    Next
                Next
            Else
                ' otherwise spin through map markers using icon names to find a display name
                allGroups = doc.MapMarkerGroups
                ' Debug.Clear
                For Each aGroup In allGroups
                    ' Debug.Print "======"
                    ' Debug.Print "Group " + aGroup.Name + " contains:"
                    For Each aMapMarker In aGroup
                        ' Debug.Print aMapMarker.Label + " == " + aMapMarker.Icon.Name
                        If StrComp(aMapMarker.Icon.Name, ic.Name) = 0 Then
                            result = aMapMarker.Label   ' found a display name so exit
                            GoTo ExitNow
                        End If
                    Next
                Next
            End If

        Catch
        End Try

ExitNow:
        GetIconDisplayName = result
    End Function

    Sub AddIconfromDisplayName(ByRef m_app As mm.Application, ByRef doc As mm.Document, ByRef t As mm.Topic, ByRef icstringold As String)
        Dim allGroups As mm.MapMarkerGroups
        Dim aGroup As mm.MapMarkerGroup
        Dim aMapMarker As mm.MapMarker
        Dim combinedstring As String
        Dim groupstring As String
        Dim iconstring As String
        combinedstring = Mid(icstringold, 6, Len(icstringold)) 'trim off leading icon:
        groupstring = Mid(combinedstring, 1, InStr(combinedstring, ":") - 1)
        iconstring = Mid(combinedstring, InStr(combinedstring, ":") + 1, Len(combinedstring))
        If InStr(iconstring, "StockIcon-") > 0 Then
            t.AllIcons.AddStockIcon(CType(Val(Mid(iconstring, InStr(iconstring, "-") + 1, Len(iconstring))), Integer))
        Else
            allGroups = doc.MapMarkerGroups
            For Each aGroup In allGroups
                If (aGroup.Name = groupstring) And ((aGroup.Type = mm.MmMapMarkerGroupType.mmMapMarkerGroupTypeIcon) Or (aGroup.Type = mm.MmMapMarkerGroupType.mmMapMarkerGroupTypeSingleIcon)) Then
                    For Each aMapMarker In aGroup
                        If StrComp(aMapMarker.Label, iconstring) = 0 Then
                            If (aMapMarker.Icon.Type = mm.MmIconType.mmIconTypeStock) Then
                                'Debug.Print("found a stock icon for " & icstring)
                                Try
                                    t.AllIcons.AddStockIcon(aMapMarker.Icon.StockIcon)
                                Catch
                                    MsgBox("Can not add " & iconstring)
                                End Try
                                Exit Sub
                            Else
                                If (aMapMarker.Icon.Type = mm.MmIconType.mmIconTypeCustom) Then
                                    'Try
                                    Dim output As String
                                    output = FindCustomIconFileNameFromSignature(m_app, aMapMarker.Icon.CustomIconSignature, "")
                                    If Len(output) > 0 Then t.AllIcons.AddCustomIcon(output)
                                    'Catch
                                    'End Try

                                    Exit Sub
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
    End Sub
    Function getrange(ByRef sheet As excel.Worksheet, ByVal row As Integer, ByVal col As Integer) As excel.Range
        getrange = sheet.Range(convertaddress(row, col))
    End Function
    Function GetcustomIconFromDisplayName(ByRef doc As mm.Document, ByRef icstring As String) As mm.MapMarkerIcon
        Dim allGroups As mm.MapMarkerGroups
        Dim aGroup As mm.MapMarkerGroup
        Dim aMapMarker As mm.MapMarker
        Dim result As mm.MapMarkerIcon
        result = Nothing
        Try
            If InStr(icstring, "Custom") > 0 Then
                allGroups = doc.MapMarkerGroups
                For Each aGroup In allGroups
                    For Each aMapMarker In aGroup
                        If (aMapMarker.Icon.Type = mm.MmIconType.mmIconTypeCustom) Then ' only mmIconTypeCustom types have CustomIconSignature properties
                            If StrComp(aMapMarker.Label, icstring) = 0 Then
                                result = aMapMarker.Icon   ' found a display name so exit
                                GoTo ExitNow
                            End If
                        End If
                    Next
                Next
            Else
                result = Nothing
            End If

        Catch
        End Try

ExitNow:
        GetcustomIconFromDisplayName = result
    End Function


End Module
