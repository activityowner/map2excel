Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
'new version 

Module Module1
    Public hasimages As Boolean = False
    Public Const startdatecol = 1
    Public Const duedatecol = 2
    Public Const prioritycol = 3
    Public Const whocol = 4
    Public Const pctdonecol = 5
    Public Const imagecol = 6
    Public notescol As Integer
    Public Const titlerow = 1
    Public Const titlecol = 1
    Public headerrow As Integer = 2


    Sub Map2Excel(ByRef m_app As Mindjet.MindManager.Interop.Application, ByRef fillin As Boolean, ByRef addextended As Boolean)
        Dim sheet As excel.Worksheet
        Dim Parent As mm.Topic
        Dim t As mm.Topic
        Dim rownum As Integer
        Dim maxcolnum As Integer
        Dim topicstart As Integer
        Dim i As Integer
        Dim iconstrings As ArrayList
        Dim custompropertystrings As ArrayList
        Dim tagstrings As ArrayList
        Dim f As New ImportInProgress
        If Not donated() Then
            MsgBox("Activation required")
            Exit Sub
        End If
        Parent = getparent(m_app)
        sheet = openexcelsheet()
        'Dim sheet2 As excel.Worksheet
        Dim Stopwatch = New Stopwatch()
        Stopwatch.Start()
        'Debug.Print("about to write to excel")
        'sheet2 = sheet.Application.Sheets.Item(2)
        'sheet2.Activate()
        'sheet2.Cells(1, 1) = "Exporting Map. This may take a few minutes"

        f.Show()
        f.Text = "Export to Excel in Progress"
        f.Activate()

        iconstrings = IconStringList(m_app.ActiveDocument)
        tagstrings = tagstringList(m_app.ActiveDocument)
        custompropertystrings = custompropertystringlist(m_app.ActiveDocument)


        'WRITE HEADER-------------------------------------------------
        With sheet
            If fillin Then
                .Cells(titlerow, titlecol) = "Table:" & Parent.Text
            Else
                .Cells(titlerow, titlecol) = Parent.Text
            End If

            .Range(convertaddress(titlerow, titlecol)).WrapText = False
            .Range(convertaddress(titlerow, titlecol)).Font.Bold = True
        End With

        If addextended Then
            With sheet
                .Cells(headerrow, startdatecol) = "Start"
                setcolwidth(sheet, startdatecol, 10)
                .Cells(headerrow, duedatecol) = "Due"
                setcolwidth(sheet, duedatecol, 10)
                .Cells(headerrow, prioritycol) = "P"
                .Cells(headerrow, whocol) = "Who"
                .Cells(headerrow, pctdonecol) = "%"
                '.Cells(headerrow, tagscol) = "Tags"
                .Cells(headerrow, imagecol) = "Image"
            End With

            For i = 0 To iconstrings.Count - 1
                sheet.Cells(headerrow, imagecol + i + 1) = "icon:" & iconstrings.Item(i)
                setcolwidth(sheet, imagecol + i + 1, 2)
            Next
            Debug.Print("writing tag header")
            Debug.Print(tagstrings.Count)
            For i = 0 To tagstrings.Count - 1
                sheet.Cells(headerrow, imagecol + iconstrings.Count + i + 1) = "tag:" & tagstrings.Item(i)
                setcolwidth(sheet, imagecol + iconstrings.Count + i + 1, 2)
            Next
            Debug.Print("writing custom property header")
            For i = 0 To custompropertystrings.Count - 1
                sheet.Cells(headerrow, imagecol + iconstrings.Count + tagstrings.Count + i + 1) = "custom:" & custompropertystrings.Item(i)
                setcolwidth(sheet, imagecol + iconstrings.Count + tagstrings.Count + i + 1, 2)
            Next

            notescol = imagecol + iconstrings.Count + tagstrings.Count + custompropertystrings.Count + 1

            If getmtckey("options", "notesincomments") = "0" Then
                sheet.Cells(headerrow, notescol) = "Notes"
                setcolwidth(sheet, notescol, 100)
            Else
                sheet.Cells(headerrow, notescol) = "Notes in Comments"
                setcolwidth(sheet, notescol, 2)
            End If
            topicstart = notescol + 1
        Else
            topicstart = 1
        End If

        maxcolnum = 1
        Dim outlinestr As String
        Dim outlinenum As Integer
        outlinenum = 1
        rownum = headerrow + 1
        'MAIN EXPORT LOOP-----------------------------------------------------------------------------------

        f.ProgressLabel.Text = "0%"
        Dim progressdenominator As Integer
        progressdenominator = Parent.AllSubTopics.Count
        For Each t In Parent.AllSubTopics
            If t.IsVisible Or Not (getmtckey("options", "limittovisible") = "1") Then
                outlinestr = Trim(Str(outlinenum)) & "."
                ExportTopic2Excel(Parent, t, rownum, topicstart, sheet, fillin, addextended, maxcolnum, iconstrings, tagstrings, custompropertystrings, outlinestr)
                f.ProgressLabel.Text = Str(System.Math.Round((outlinenum / progressdenominator) * 100)) & "% complete"
                outlinenum = outlinenum + 1
            End If

        Next

        'ADD LEVEL HEADER------------------------------------------------------------------------------------
        For i = topicstart To maxcolnum
            sheet.Cells(headerrow, i) = "L" & Trim(Str(i - topicstart + 1))
        Next

        For i = 1 To maxcolnum
            sheet.Range(convertaddress(headerrow, i)).Orientation = 90
        Next

        rownum = findlastrow(sheet)
        Dim maxwidth As Integer
        For i = 1 To imagecol
            maxwidth = 2
            For j = headerrow + 1 To rownum
                If Len(celltext(sheet, j, i)) > maxwidth Then maxwidth = Len(celltext(sheet, j, i))
            Next
            setcolwidth(sheet, i, maxwidth)
        Next
        'narrowupcolumns(sheet, maxcolnum, rownum)
        Parent = Nothing
        t = Nothing
        f.Close()
        sheet.Activate()
        sheet.Cells(1, 1).Select()
        MsgBox("Export Complete (" & System.Math.Round(Stopwatch.ElapsedMilliseconds / 1000) & " seconds)")
        'If rownum > 30 And (Not donated()) Then System.Diagnostics.Process.Start("http://www.activityowner.com/map2excel-trial-period/")
        'sheet2.Cells(1, 1) = ""
        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet2)
        sheet = Nothing
        'sheet2 = Nothing
        f = Nothing

    End Sub
    Function donated() As Boolean
        donated = True
        'If (getmtckey("options", "key") = "AO20092012") Or (Left(getmtckey("options", "key"), 6) = "AO2016" And Right(getmtckey("options", "key"), 1) = "6") Then
        'donated = True
        'ElseIf intrial() Then
        'donated = True
        'Else
        'donated = True
        'End If
    End Function

    Function intrial() As Boolean
        Dim firstRunDate As String
        firstRunDate = getmtckey("options", "frmtc")

        If firstRunDate = Nothing Then
            firstRunDate = DateString()
            setmtckey("options", "frmtc", firstRunDate)
        End If
        Try
            If (Now - Date.Parse(firstRunDate)).Days > 7 Then
                intrial = True
                'MsgBox("Trial Expired")
                'intrial = False
            Else
                intrial = True
                MsgBox(Str(7 - (Now - Date.Parse(firstRunDate)).Days) & " days left in trial")
            End If
        Catch
            MsgBox("Licence Error: Email activityowner@gmail.com for support")
        End Try

    End Function

End Module
