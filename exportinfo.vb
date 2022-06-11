Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Module exportinfo
    Sub ExportTopic2Excel(ByRef parent As mm.Topic, ByRef t As mm.Topic, ByRef rownum As Integer, ByRef colnum As Integer, ByRef sheet As excel.Worksheet, ByRef fillin As Boolean, ByRef addextended As Boolean, ByRef maxcolnum As Integer, ByRef iconstrings As ArrayList, ByRef tagstrings As ArrayList, ByVal custompropertystrings As ArrayList, ByVal outlinestr As String)
        Dim st As mm.Topic
        Dim p As mm.Topic
        Dim i As Integer
        Dim k As Integer
        Dim tags As String
        Dim found As Boolean
        Dim imagefilename As String
        Dim tempstate As Boolean
        Dim rng As excel.Range
        Dim topictext As String

        rng = sheet.Range(convertaddress(rownum, colnum))
        'adjustcolwidth(sheet, t, maxcolnum, colnum, rng)

        If getmtckey("options", "addoutlinenumbers") = "1" Then
            topictext = outlinestr & " " & t.Text
        Else
            topictext = t.Text
        End If

        rng.Value = topictext 'put text in cell
        If Not getmtckey("options", "wraptext") = "1" Then rng.WrapText = False

        If getmtckey("options", "addtopichyperlinks") = "1" Then AddTopicLink(rng, t, topictext)
        If getmtckey("options", "addhyperlinks") = "1" Then addhyperlinkentry(sheet, t, rng, topictext)
        putnotesincomment(rng, t.Notes.Text)
        copyfontcolortoExcel(t, rng)

        If fillin Then
            p = t.ParentTopic
            k = colnum - 1
            While Not (p Is parent)
                If getmtckey("options", "addoutlinenumbers") = "1" Then
                    topictext = outlinestr & " " & p.Text
                Else
                    topictext = p.Text
                End If
                rng = sheet.Range(convertaddress(rownum, k))
                rng.Value = p.Text
                copyfontcolortoExcel(p, rng)
                If getmtckey("options", "addtopichyperlinks") = "1" Then AddTopicLink(rng, p, topictext)
                If getmtckey("options", "addhyperlinks") = "1" Then addhyperlinkentry(sheet, p, rng, topictext)
                putnotesincomment(rng, p.Notes.Text)
                k = k - 1
                p = p.ParentTopic
            End While
        End If


        'If addextended And (t.AllSubTopics.Count = 0 Or Not fillin) Then
        If addextended Then
            If Not isdate0(t.Task.StartDate) Then sheet.Cells(rownum, startdatecol) = t.Task.StartDate
            If Not isdate0(t.Task.DueDate) Then sheet.Cells(rownum, duedatecol) = t.Task.DueDate
            If t.Task.Priority > 0 Then sheet.Cells(rownum, prioritycol) = t.Task.Priority
            If t.Task.Complete > -1 Then sheet.Cells(rownum, pctdonecol) = t.Task.Complete

            If Len(t.Task.Resources) > 0 Then
                sheet.Cells(rownum, whocol) = t.Task.Resources
                If Len(t.Task.Resources) > CType(sheet.Range(sheet.Cells(1, whocol), sheet.Cells(1, whocol)).ColumnWidth, Integer) Then
                    sheet.Range(sheet.Cells(1, whocol), sheet.Cells(1, whocol)).ColumnWidth = Len(t.Task.Resources)
                End If
            End If

            If t.TextLabels.Count > 0 Then
                tags = ""
                For i = 1 To t.TextLabels.Count
                    tags = tags & t.TextLabels.Item(i).Name
                    If i < t.TextLabels.Count Then tags = tags & ","
                Next
                'sheet.Cells(rownum, 6) = tags
                'If Len(tags) > CType(sheet.Range(convertaddress(1, 7)).ColumnWidth, Integer) Then
                ' sheet.Range(convertaddress(1, 7)).ColumnWidth = Len(tags)
                'End If
            End If

            rng = sheet.Range(convertaddress(rownum, imagecol))

            imagefilename = System.IO.Path.GetTempFileName
            If t.HasImage Then
                If getmtckey("options", "addimagetocell") = "1" Then
                    t.Image.Save(imagefilename, Mindjet.MindManager.Interop.MmGraphicType.mmGraphicTypeBmp)
                    setcolwidth(sheet, imagecol, 10)
                    setrowheight(sheet, rownum, 50)
                    InsertPictureInRange(imagefilename, rng, sheet)
                End If
                If getmtckey("options", "addimagetocomment") = "1" Then
                    t.Image.Save(imagefilename, Mindjet.MindManager.Interop.MmGraphicType.mmGraphicTypeBmp)
                    InsertPictureInComment(imagefilename, rng, sheet)
                End If
            Else
                If t.Attachments.Count > 0 Then
                    If InStr(LCase(t.Attachments.Item(1).FileName), "jpeg") > 0 Or InStr(LCase(t.Attachments.Item(1).FileName), "jpg") > 0 Then
                        t.Attachments.Item(1).SaveAs(imagefilename)
                        If getmtckey("options", "addimagetocomment") = "1" Then
                            InsertPictureInComment(imagefilename, rng, sheet)
                        Else
                            InsertPictureInRange(imagefilename, rng, sheet)
                        End If
                    End If
                Else
                    If t.HasHyperlink Then
                        If InStr(LCase(t.Hyperlink.Address), "jpg") > 0 Or InStr(LCase(t.Hyperlink.Address), "jpeg") > 0 Then
                            If getmtckey("options", "addimagetocomment") = "1" Then
                                tempstate = t.Hyperlink.Absolute
                                t.Hyperlink.Absolute = True
                                InsertPictureInComment(t.Hyperlink.Address, rng, sheet)
                                t.Hyperlink.Absolute = tempstate
                            Else
                                tempstate = t.Hyperlink.Absolute
                                t.Hyperlink.Absolute = True
                                InsertPictureInRange(t.Hyperlink.Address, rng, sheet)
                                t.Hyperlink.Absolute = tempstate
                            End If
                        End If
                    End If
                End If
            End If

            For i = 0 To iconstrings.Count - 1
                If hasicon(t.Document, t, iconstrings.Item(i).ToString) Then
                    sheet.Cells(rownum, i + imagecol + 1) = 1
                End If
            Next
            For i = 0 To tagstrings.Count - 1
                If t.TextLabels.Count > 0 Then
                    For j = 1 To t.TextLabels.Count
                        If StrComp(tagstrings.Item(i).ToString, t.TextLabels.Item(j).Name) = 0 Then
                            sheet.Cells(rownum, i + imagecol + iconstrings.Count + 1) = 1
                        End If
                    Next

                End If
            Next
            For i = 0 To custompropertystrings.Count - 1
                If hascustomproperty(t, custompropertystrings.Item(i).ToString) Then
                    sheet.Cells(rownum, i + imagecol + iconstrings.Count + tagstrings.Count + 1) = getcustompropertyvalue(t, custompropertystrings.Item(i).ToString)
                End If
            Next



            If Len(t.Notes.Text) > 0 Then
                If getmtckey("options", "notesincomments") = "1" Then
                    Try
                        rng = sheet.Range(convertaddress(rownum, colnum))
                        With rng
                            .AddComment()
                            .Comment.Text(Text:=t.Notes.Text)
                            .Comment.Shape.ScaleHeight(5, 0)
                            .Comment.Shape.ScaleWidth(4, 0)
                        End With
                    Catch
                    End Try
                Else
                    sheet.Cells(rownum, imagecol + iconstrings.Count + 1) = t.Notes.Text
                End If
            End If

        End If

        If (t.AllSubTopics.Count = 0) Then
            If colnum > maxcolnum Then maxcolnum = colnum
            rownum = rownum + 1
        Else
            colnum = colnum + 1
            If colnum > maxcolnum Then maxcolnum = colnum
            found = False
            i = 1
            For Each st In t.AllSubTopics
                If st.IsVisible Or Not getmtckey("options", "limittovisible") = "1" Then
                    If Not found Then 'If Not fillin And Not found Then
                        found = True
                        rownum = rownum + 1 ' outline
                    End If
                    ExportTopic2Excel(parent, st, rownum, colnum, sheet, fillin, addextended, maxcolnum, iconstrings, tagstrings, custompropertystrings, outlinestr & Trim(Str(i)) & ".")
                    i = i + 1
                End If
            Next
            colnum = colnum - 1
        End If
        st = Nothing
    End Sub
End Module
