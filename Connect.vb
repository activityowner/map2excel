imports Extensibility
imports System.Runtime.InteropServices

#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the Map2ExcelSetup project, 
' right click the project in the Solution Explorer, then choose install.
#End Region

<GuidAttribute("09EA00D7-A7F1-4303-9858-8873B559D0F9"), ProgIdAttribute("Map2Excel.Connect")> _
Public Class Connect

    Implements Extensibility.IDTExtensibility2

    Private m_applicationObject As Mindjet.MindManager.Interop.Application
    Private o_addInInstance As Object

    Dim myRibbon As MindManagerRibbon.M2ERibbonGroup


    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete
    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection

        m_applicationObject = Nothing
        o_addInInstance = Nothing
        myRibbon.Dispose()

        System.GC.Collect()

        'MsgBox("Goodbye MindManager Debvelopment Universe!")

    End Sub

    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As Extensibility.ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection

        m_applicationObject = CType(application, Mindjet.MindManager.Interop.Application)
        o_addInInstance = addInInst

        myRibbon = New MindManagerRibbon.M2ERibbonGroup(m_applicationObject)

        'MsgBox("Hello MindManager Development Universe!")

    End Sub
End Class