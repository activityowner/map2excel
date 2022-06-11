Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Module excel2mapcode
    Sub excellinks2map(ByRef m_app As Mindjet.MindManager.Interop.Application)
        Dim newmap As mm.Document
        Dim sheet As excel.Worksheet
        Debug.Print("Starting Excellinks2Map===========================")
        Dim statusmap As mm.Document

        If Len(getmtckey("options", "templatefile")) > 0 Then
            Try
                newmap = m_app.AllDocuments.AddFromTemplate(getmtckey("options", "templatefile"))
            Catch
                MsgBox("Error loading template file")
            End Try
        Else
            newmap = m_app.AllDocuments.Add()
        End If
        Dim p As mm.Topic
        sheet = getsheetforimport("d:\\test.xlsx")
        p = newmap.CentralTopic
        Dim lastrow As Integer
        lastrow = findlastrow(sheet)
        Dim importdialog As New ImportInProgress
        Dim t As mm.Topic
        For i = 1 To lastrow
            t = st_create(p.Document, p, celltext(sheet, i, 1))
            t.CreateHyperlink(celltext(sheet, i, 2))
        Next
        newmap.Activate()
        newmap = Nothing
        If sheet.Application.ActiveWorkbook.Saved = True Then
            sheet.Application.ActiveWorkbook.Close()
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        sheet = Nothing
    End Sub

    Sub excel2map(ByRef m_app As Mindjet.MindManager.Interop.Application)
        Dim newmap As mm.Document
        Dim sheet As excel.Worksheet
        Debug.Print("Starting Excel2Map===========================")
        Dim statusmap As mm.Document

        If Len(getmtckey("options", "templatefile")) > 0 Then
            Try
                newmap = m_app.AllDocuments.AddFromTemplate(getmtckey("options", "templatefile"))
            Catch
                MsgBox("Error loading template file")
            End Try
        Else
            newmap = m_app.AllDocuments.Add()
        End If


        Dim row As Integer
        Dim startrow As Integer
        Dim col As Integer
        Dim p As mm.Topic

        sheet = getsheetforimport("d:\\test.xlsx")
        statusmap = m_app.AllDocuments.Add(True)
        statusmap.CentralTopic.Text = "Importing excel sheet. This may take a few minutes"
        Dim startcol As Integer
        startcol = findstartcol(sheet, headerrow)
        If startcol = 1 Then
            startrow = headerrow
        Else
            startrow = headerrow + 1
        End If

        If startcol = 1 Then
            If Not StrComp(celltext(sheet, headerrow, 1), "L1") = 0 Then
                row = 2
            End If
        End If

        p = newmap.CentralTopic
        Dim importastable As Boolean
        importastable = InStr(celltext(sheet, 1, 1), "Table:") = 1
        If importastable Then
            p.Text = Replace(celltext(sheet, 1, 1), "Table:", "")
        Else
            p.Text = celltext(sheet, 1, 1)
        End If


        col = startcol
        Dim lastrow As Integer
        lastrow = findlastrow(sheet)
        Dim Stopwatch = New Stopwatch()
        Stopwatch.Start()
        Dim importdialog As New ImportInProgress
        If Not importastable Then ' MsgBox("Import sheet as outline? (no=import as table)", vbYesNo) = vbYes Then
            importdialog.Show()
            importdialog.Activate()
            getkids(m_app, p, sheet, startrow, col, importdialog, lastrow, startcol)
        Else
            For i = startrow To lastrow
                importdialog.Show()
                importdialog.Activate()
                gettablerow(m_app, p, sheet, i, col, importdialog, lastrow, startcol)
            Next
        End If

        importdialog.Close()
        Stopwatch.Stop()
        statusmap.Close()
        newmap.Activate()
        Debug.Print(Str(Stopwatch.ElapsedMilliseconds / 1000))
        newmap = Nothing
        If sheet.Application.ActiveWorkbook.Saved = True Then
            sheet.Application.ActiveWorkbook.Close()
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        sheet = Nothing
        MsgBox("Import Complete (" & System.Math.Round(Stopwatch.ElapsedMilliseconds / 1000) & " seconds)")
        'If lastrow > 30 And (Not donated()) Then System.Diagnostics.Process.Start("http://www.activityowner.com/map2excel-trial-period/")
    End Sub

    Function getkids(ByRef m_app As mm.Application, ByRef p As mm.Topic, ByRef sheet As excel.Worksheet, ByVal row As Integer, ByVal col As Integer, ByRef importdialog As ImportInProgress, ByVal lastrow As Integer, ByVal startcol As Integer
                     ) As Integer
        Dim t As mm.Topic
        'add self to parent
        importdialog.ProgressLabel.Text = Str(System.Math.Round((row - headerrow) / (lastrow - headerrow) * 100)) & "% complete: " & celltext(sheet, row, col)

        t = st_create(p.Document, p, celltext(sheet, row, col))

        If Not startcol = 1 Then addattributes(m_app, t, sheet, row, col)

        If Len(celltext(sheet, row + 1, col + 1)) > 0 Then
            row = getkids(m_app, t, sheet, row + 1, col + 1, importdialog, lastrow, startcol) 'get subtopic
        End If

        If Len(celltext(sheet, row + 1, col)) > 0 Then
            row = getkids(m_app, p, sheet, row + 1, col, importdialog, lastrow, startcol)
        End If
        getkids = row
    End Function

    Function gettablerow(ByRef m_app As mm.Application, ByRef p As mm.Topic, ByRef sheet As excel.Worksheet, ByVal row As Integer, ByVal col As Integer, ByRef importdialog As ImportInProgress, ByVal lastrow As Integer, ByVal startcol As Integer
                    ) As Integer
        Dim t As mm.Topic
        'add self to parent
        importdialog.ProgressLabel.Text = Str(System.Math.Round((row - headerrow) / (lastrow - headerrow) * 100)) & "% complete" & celltext(sheet, row, col)

        t = st_create(p.Document, p, celltext(sheet, row, col))


        If Len(celltext(sheet, row, col + 1)) > 0 Then
            row = gettablerow(m_app, t, sheet, row, col + 1, importdialog, lastrow, startcol) 'get subtopic
        Else
            If Not startcol = 1 Then addattributes(m_app, t, sheet, row, col) 'WHAT ABOUT THISSSSSSSSSS
        End If

    End Function
    Function getmapforimport() As String
        getmapforimport = Nothing
        Try
            Dim fd As OpenFileDialog = New OpenFileDialog()
            fd.Title = "Open Template Map with Map Marker Groups"
            fd.InitialDirectory = "C:\"
            fd.Filter = "Mindjet Template files (*.mmat)|*.mmat"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True
            If fd.ShowDialog() = DialogResult.OK Then
                getmapforimport = fd.FileName
            Else
                getmapforimport = Nothing
            End If
        Catch
            MsgBox("error opening file")
        End Try
    End Function

    Function getsheetforimport(ByVal defaultfile As String) As excel.Worksheet
        Dim excelworkbook As excel.Workbook
        Dim excelapp = New excel.Application
        getsheetforimport = Nothing
        Try
            Dim fd As OpenFileDialog = New OpenFileDialog()
            fd.Title = "Open Excel File to be Imported to Mindjet"
            fd.InitialDirectory = "C:\"
            fd.Filter = "Excel files (*.xls*)|*.xls*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True
            If fd.ShowDialog() = DialogResult.OK Then
                defaultfile = fd.FileName
                excelworkbook = excelapp.Workbooks.Open(defaultfile)
                excelapp.Visible = True
                getsheetforimport = CType(excelworkbook.Sheets.Item(1), excel.Worksheet)
                excelworkbook.Activate()
            Else
                getsheetforimport = Nothing
            End If
        Catch
            MsgBox("error opening file")
        End Try
        excelapp = Nothing
        excelworkbook = Nothing
    End Function


    Function findlastrow(ByRef objworksheet As excel.Worksheet) As Integer
        'from:http://blogs.technet.com/b/heyscriptingguy/archive/2006/02/15/how-can-i-determine-the-last-row-in-an-excel-spreadsheet.aspx
        Const xlCellTypeLastCell = 11
        Dim objRange As excel.Range
        objworksheet.Activate()
        objRange = objworksheet.UsedRange
        objRange.SpecialCells(xlCellTypeLastCell).Activate()
        findlastrow = objworksheet.Application.ActiveCell.Row
    End Function

    Sub addattributes(ByRef m_app As mm.Application, ByRef t As mm.Topic, ByRef sheet As excel.Worksheet, ByRef row As Integer, ByRef col As Integer)
        If Len(celltext(sheet, row, startdatecol)) > 0 Then t.Task.StartDate = celltext(sheet, row, startdatecol)
        If Len(celltext(sheet, row, duedatecol)) > 0 Then t.Task.DueDate = celltext(sheet, row, duedatecol)
        If Len(celltext(sheet, row, prioritycol)) > 0 Then t.Task.Priority = celltext(sheet, row, prioritycol)
        If Len(celltext(sheet, row, whocol)) > 0 Then t.Task.Resources = celltext(sheet, row, whocol)
        If Len(celltext(sheet, row, pctdonecol)) > 0 Then t.Task.Complete = celltext(sheet, row, pctdonecol)

        If Len(getcellhyperlink(sheet, row, col)) > 0 Then t.CreateHyperlink(getcellhyperlink(sheet, row, col))

   
        'Notes
        notescol = findstartcol(sheet, headerrow) - 1
        If getmtckey("options", "notesincomments") = "0" Then
            If Len(celltext(sheet, row, notescol)) > 0 Then
                t.Notes.Text = celltext(sheet, row, notescol)
                t.Notes.Commit()
            End If
        Else
            Try
                If Len(getrange(sheet, row, col).Comment.Text) > 0 Then
                    t.Notes.Text = getrange(sheet, row, col).Comment.Text
                    t.Notes.Commit()
                End If
            Catch
            End Try
        End If


        Dim i As Integer
        'Icons
        For i = imagecol + 1 To notescol - 1
            If InStr(celltext(sheet, headerrow, i), "icon:") > 0 Then
                If Len(celltext(sheet, row, i)) > 0 Then
                    'AddIconfromDisplayName(m_app, t.Document, t, Replace(celltext(sheet, headerrow, i), "icon:", ""))
                    AddIconfromDisplayName(m_app, t.Document, t, celltext(sheet, headerrow, i))
                End If
            End If
        Next


        For i = imagecol + 1 To notescol - 1
            If InStr(celltext(sheet, headerrow, i), "tag:") > 0 Then
                If Len(celltext(sheet, row, i)) > 0 Then
                    t.TextLabels.AddTextLabel(Replace(celltext(sheet, headerrow, i), "tag:", ""))
                End If
            End If
        Next

        'Tags

        'If Len(celltext(sheet, row, tagscol)) > 0 Then addtags(t, celltext(sheet, row, tagscol))

        'Image
        '    sheet.Cells(headerrow, imagecol) = "Image"



    End Sub
    Sub addtags(ByRef t As mm.Topic, ByVal tagstrings As ArrayList, ByVal tagbinary As ArrayList)
        Dim i As Integer
        Dim s As String
        Dim j As Integer

        For i = 0 To tagbinary.Count
            If tagbinary(i) = 1 Then t.TextLabels.AddTextLabel(tagstrings(i))
        Next

        'Process,In-tray*
        'While Len(tags) > 0
        'If InStr(tags, ",") > 0 Then
        ' t.TextLabels.AddTextLabel(Mid(tags, 1, InStr(tags, ",") - 1))
        ' tags = Mid(tags, InStr(tags, ",") + 1, Len(tags))
        ' Else
        ' t.TextLabels.AddTextLabel(tags)
        ' tags = ""
        'End If
        'End While
    End Sub
    Function findstartcol(ByRef sheet As excel.Worksheet, ByVal headerrow As Integer) As Integer
        findstartcol = 1
        Const findstartcolmax = 256
        Try
            While Not (celltext(sheet, headerrow, findstartcol) = "L1" Or celltext(sheet, headerrow, findstartcol) = "L 1") And findstartcol < findstartcolmax
                findstartcol = findstartcol + 1
            End While
        Catch
            findstartcol = 1
        End Try
        If findstartcol = findstartcolmax Then
            If Not celltext(sheet, headerrow, startdatecol) = "Start" Then
                findstartcol = 1
            Else
                MsgBox("Error:  L1 column not found")
                findstartcol = -1
            End If
        End If
    End Function
    Function celltext(ByRef sheet As excel.Worksheet, ByVal row As Integer, ByRef col As Integer) As String
        celltext = CType(sheet.Range(convertaddress(row, col)).Text, String)
    End Function

    Function getcellhyperlink(ByRef sheet As excel.Worksheet, ByVal row As Integer, ByVal col As Integer) As String
        Dim rng As excel.Range
        Debug.Print("looking for hyperlinks")
        rng = getrange(sheet, row, col)
        If rng.Hyperlinks.Count > 0 Then
            getcellhyperlink = rng.Hyperlinks.Item(1).Address
        ElseIf rng.HasFormula Then
            Debug.Print("has formula" & rng.Formula.ToString)
            If InStr(rng.Formula.ToString, "=HYPERLINK(") > 0 Then
                getcellhyperlink = Replace(rng.Formula.ToString, "=HYPERLINK(" & Chr(34), "")
                Debug.Print(getcellhyperlink)
                getcellhyperlink = Mid(getcellhyperlink, 1, InStr(getcellhyperlink, Chr(34)) - 1)
                Debug.Print(getcellhyperlink)
            Else
                getcellhyperlink = ""
            End If
        Else
            getcellhyperlink = ""
        End If
        If InStr(getcellhyperlink, "mj-map") > 0 Then getcellhyperlink = "" 'don't add a topic link back?   but what if different map
    End Function
End Module
