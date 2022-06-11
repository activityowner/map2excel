
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports mm = Mindjet.MindManager.Interop
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Imports excel = Microsoft.Office.Interop.Excel
'bugs
' can't close map
Namespace MindManagerRibbon

    Class M2ERibbonGroup

        Implements IDisposable

#Region "Variables"
        Private m_app As mm.Application
        Private WithEvents myCommand1 As mm.Command
        Private WithEvents myCommand2 As mm.Command
        Private WithEvents myCommand3 As mm.Command
        Private WithEvents myCommand4 As mm.Command
        Private WithEvents myCommand5 As mm.Command
        Private WithEvents myCommand6 As mm.Command
        Private WithEvents myCommand7 As mm.Command
        Private WithEvents myCommand8 As mm.Command
        Private WithEvents myCommand9 As mm.Command
        Private WithEvents mycommand10 As mm.Command
        Public usedit As Boolean


#End Region

        Public Sub New(ByVal app As Mindjet.MindManager.Interop.Application)
            Try
                m_app = app
                usedit = False
                'Creates the Ribbon
                Dim newRibbontab As Mindjet.MindManager.Interop.ribbonTab = Ribbons.CreateRibbon(m_app, "Map2Excel", "urn:M2EG.Tab")

                'Creates the Ribbon Groups

                Dim newribbongroup4 As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(newRibbontab, "Information", "urn:M2EG.Group3")
                Dim newribbongroup3 As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(newRibbontab, "Extras", "urn:M2EG.Group2")
                Dim newribbongroup1 As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(newRibbontab, "Map2Excel Export", "urn:M2EG.Group1")
                Dim newribbongroup5 As Mindjet.MindManager.Interop.RibbonGroup = Ribbons.CreateGroupTab(newRibbontab, "Excel2Map", "urn:M2EG.Group4")

                'Creates the Ribbon Group Commands
                myCommand1 = m_app.Commands.Add("Map2Excel.Connect", "OutlineSimple")
                myCommand2 = m_app.Commands.Add("Map2Excel.Connect", "OutlineWithDetails")
                myCommand3 = m_app.Commands.Add("Map2Excel.Connect", "TableSimple")
                myCommand4 = m_app.Commands.Add("Map2Excel.Connect", "TablewithDetails")
                myCommand5 = m_app.Commands.Add("Map2Excel.Connect", "Map2Outline")
                myCommand8 = m_app.Commands.Add("Map2Excel.Connect", "Map2Table")
                myCommand9 = m_app.Commands.Add("Map2Excel.Connect", "Options")
                mycommand10 = m_app.Commands.Add("Map2Excel.Connect", "Excel2Map")
                myCommand1.BasicCommand = True
                myCommand2.BasicCommand = True
                myCommand3.BasicCommand = True
                myCommand4.BasicCommand = True
                myCommand5.BasicCommand = True
                myCommand9.BasicCommand = True
                mycommand10.BasicCommand = True

                myCommand1.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\excel.jpg"
                myCommand2.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\excel.jpg"
                myCommand3.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\excel.jpg"
                myCommand4.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\excel.jpg"
                myCommand5.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\ao.jpg"
                myCommand8.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\ao.jpg"
                myCommand9.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\ao.jpg"
                mycommand10.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\excel.jpg"
                myCommand1.ToolTip = ("Simple Outline Export" + (Chr(10)) + "")
                myCommand2.ToolTip = ("Export Outline with Details" + (Chr(10)) + "includes priority, resources, etc.")
                myCommand3.ToolTip = ("Simple Table" + (Chr(10)) + "")
                myCommand4.ToolTip = ("Table with Details" + (Chr(10)) + "includes priority, resources, etc.")
                myCommand5.ToolTip = ("Export Outline to text file" + (Chr(10)) + "")
                myCommand8.ToolTip = ("Export Map to an html table" + (Chr(10)) + "")
                mycommand10.ToolTip = ("Excel to Map" + (Chr(10)) + "excel to map")

                myCommand1.Caption = "Outline"
                myCommand2.Caption = "Outline with Details" ' Outline With Details"
                myCommand3.Caption = "Table"
                myCommand4.Caption = "Table With Details"
                myCommand5.Caption = "Text File Outline"
                myCommand8.Caption = "HTML Table"

                myCommand6 = m_app.Commands.Add("Map2Excel.Connect", "M2EHelp")
                myCommand6.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\help.jpg"
                myCommand6.BasicCommand = True
                myCommand6.ToolTip = ("Help" + (Chr(10)) + "Help")
                myCommand6.Caption = "Help"

                myCommand7 = m_app.Commands.Add("Map2Excel.Connect", "M2EDonate")
                myCommand7.ImagePath = Environ("ProgramFiles") & "\Activityowner.com\ActivityOwner.Com Map2Excel\donate.jpg"
                myCommand7.BasicCommand = True
                myCommand7.ToolTip = ("Buy" + (Chr(10)) + "Buy")
                myCommand7.Caption = "Buy"

                myCommand9.Caption = "Options"
                mycommand10.Caption = "Import" '  Outline with Details


                newribbongroup1.GroupControls.AddButton(myCommand1, 0)
                newribbongroup1.GroupControls.AddButton(myCommand2, 0)
                newribbongroup1.GroupControls.AddButton(myCommand3, 0)
                newribbongroup1.GroupControls.AddButton(myCommand4, 0)
                newribbongroup3.GroupControls.AddButton(myCommand5, 0)
                newribbongroup3.GroupControls.AddButton(myCommand8, 0)

                newribbongroup4.GroupControls.AddButton(myCommand6, 0)
                newribbongroup4.GroupControls.AddButton(myCommand7, 0)
                newribbongroup4.GroupControls.AddButton(myCommand9, 0)
                newribbongroup5.GroupControls.AddButton(mycommand10, 0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End Sub


        Sub need_a_doc_UpdateState(ByRef pEnabled As Boolean, ByRef pChecked As Boolean) Handles myCommand1.UpdateState, myCommand2.UpdateState, myCommand3.UpdateState, myCommand4.UpdateState, myCommand5.UpdateState, myCommand6.UpdateState, myCommand8.UpdateState

            Try
                pChecked = False
                If Not m_app.ActiveDocument Is Nothing Then
                    pEnabled = True
                Else
                    pEnabled = False
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub need_nothing_UpdateState(ByRef pEnabled As Boolean, ByRef pChecked As Boolean) Handles myCommand6.UpdateState, myCommand7.UpdateState, myCommand9.UpdateState, mycommand10.UpdateState
            Try
                pChecked = False
                pEnabled = True
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        

        Sub tLock_Click1() Handles myCommand1.Click
            Try
                'visitBetaOnce()
                Map2Excel(m_app, False, False)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click2() Handles myCommand2.Click
            Try
                'visitBetaOnce()
                Map2Excel(m_app, False, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click3() Handles myCommand3.Click
            Try
                'visitBetaOnce()
                Map2Excel(m_app, True, False)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click4() Handles myCommand4.Click
            Try
                'visitBetaOnce()
                Map2Excel(m_app, True, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click5() Handles myCommand5.Click
            Try
                'visitBetaOnce()
                map2outline(m_app)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click8() Handles myCommand8.Click
            Try
                'visitBetaOnce()
                map2table(m_app)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click6() Handles myCommand6.Click
            Try
                System.Diagnostics.Process.Start("http://map2excel.com")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click7() Handles myCommand7.Click
            Try
                System.Diagnostics.Process.Start("http://activityowner.com/donate")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Sub tLock_Click10() Handles mycommand10.Click
            Try
                'visitBetaOnce()
                'excellinks2map(m_app)
                excel2map(m_app)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Sub tLock_Click9() Handles myCommand9.Click
            Try
                Dim f As New map2exceloptions
                f.CheckBox1.Checked = Not getmtckey("options", "notesincomments") = "0"
                f.limittovisible.Checked = Not getmtckey("options", "limittovisible") = "0"
                f.AddTopicHyperlinksCheckbox.Checked = Not getmtckey("options", "addtopichyperlinks") = "0"
                f.AddExternalHyperlinksCheckBox.Checked = getmtckey("options", "addhyperlinks") = "1"
                f.AddImagetoCommentCheckBox.Checked = getmtckey("options", "addimagetocomment") = "1"
                f.AddImagestoCellsCheckBox.Checked = getmtckey("options", "addimagetocell") = "1"
                f.AddOutlineNumbers.Checked = getmtckey("options", "addoutlinenumbers") = "1"
                f.licencekeybox.Text = getmtckey("options", "key")
                f.TemplateFileNameLabel.Text = getmtckey("options", "templatefile")
                f.ShowDialog()
                f.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        'Sub visitBetaOnce()
        '   If Not donated Then
        '      usedit = True
        ' End If
        'End Sub
        'Sub donatedelay()
        '   If Not donated() Then
        '      Pause(3)
        '     MsgBox("Visit the donate page to eliminate this delay and dialog box")
        'End If
        'End Sub
        Public Sub Pause(ByVal duration As Double)
            Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
            Dim dblWaitTil As Date
            Now.AddSeconds(OneSec)
            dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(duration)
            Do Until Now > dblWaitTil
                Application.DoEvents() ' Allow windows messages to be processed
            Loop
        End Sub
        Public Overloads Sub Dispose() Implements IDisposable.Dispose
            If usedit Then
                Try
                    'System.Diagnostics.Process.Start("http://wiki.activityowner.com/index.php?title=Map2Excel_Beta")
                    'If Not donated() Then System.Diagnostics.Process.Start("http://www.activityowner.com/map2excel-trial-period/")
                Catch ex As Exception

                End Try
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand1)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand2)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand3)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand4)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand5)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand6)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand7)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand8)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myCommand9)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mycommand10)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_app)

        End Sub

    End Class

End Namespace


