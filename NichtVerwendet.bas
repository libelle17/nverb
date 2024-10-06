Attribute VB_Name = "NichtVerwendet"
Sub test5()
Dim SizeW%
SizeW = 32
'Call fStSpei(HLM, "SOFTWARE\Microsoft\Fax\Inbox", "Folder", eigDatLok + "\MSFax\Inbox1")
'Call fDWSpei(HLM, "SOFTWARE\Microsoft\Fax\Inbox", "Use", 1)
'Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\Inbox", "Folder", eigDatLok + "\MSFax\Inbox2")
'Call fDWSpei(HCU, "SOFTWARE\Microsoft\Fax\Inbox", "Use", 1)
'Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "FullName", "Gerald Alexander Schade")
If 1 = 0 Then ' wirkt alles nicht
 Call fDWSpei(HLM, "SOFTWARE\Microsoft\Fax", "Dirty Days", 0)
 Call fDWSpei(HLM, "SOFTWARE\Microsoft\Fax\Devices\0000065568\{F10A5326-0261-4715-B367-2970427BBD99}", "Flags", 3)
 Call fStSpei(HLM, "SOFTWARE\Microsoft\Fax\Devices\0000065568\{F10A5326-0261-4715-B367-2970427BBD99}", "TSID", "GS 08131 616380")
 Call fStSpei(HLM, "SOFTWARE\Microsoft\Fax\Devices\0000065568\{F10A5326-0261-4715-B367-2970427BBD99}", "CSID", "GS 08131 616380")
 Call fDWSpei(HLM, "SOFTWARE\Microsoft\Fax\Devices\0000065568\{F10A5326-0261-4715-B367-2970427BBD99}", "Rings", 0)
End If
End Sub

Sub test6()
Dim objShell As Object, objFolder As Object, objfolderItem As Object, objVerb As Object
Dim colVerbs As Object, objVerbs As Object
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(Environ("windir") & "\System32")
Set objfolderItem = objFolder.ParseName("calc.exe")
 
Set colVerbs = objfolderItem.Verbs
For Each objVerb In colVerbs
   Debug.Print objVerb
Next
End Sub

Function fxs_u()
 Static Cpt$
 If Cpt = "" Then Cpt = CptName
 Dim fx As FAXCOMEXLib.FaxServer ' %windir%\system32\faxscomex.dll
 Dim fx1 As FAXCOMEXLib.FaxDocument
 Dim fx2 As FAXCONTROLLib.FaxControl
 Dim fx3 As faxcomlib.FaxTiff ' geraten aus Bezeichnung, die über oakland software viewer unter activex/com, Registry, ProgIDs steht
 Set fx = New FAXCOMEXLib.FaxServer 'CreateObject("faxcomex.faxserver")
 If Cpt = "ANMELDL1" Then Call fx.Connect("ANMELDL1") Else Call fx.Connect("ANMELDL")
 Set fx1 = New FAXCOMEXLib.FaxDocument ' CreateObject("faxcomex.faxdocument")
 Set fx3 = New faxcomlib.FaxTiff 'CreateObject("faxtiff.faxtiff.1")
 Set fx = New FAXCOMEXLib.FaxServer 'CreateObject("faxcomex.faxserver")
 If Cpt = "ANMELDL1" Then Call fx.Connect("ANMELDL1") Else Call fx.Connect("ANMELDL")
 fx.Folders.IncomingArchive.ArchiveFolder = "c:\p"
 Set fx2 = New FAXCONTROLLib.FaxControl 'CreateObject("faxcontrol.faxcontrol")
 'Set fx2 = CreateObject("faxcontrol.faxcontrol.1")
 If Not fx2.IsFaxServiceInstalled Then fx2.InstallFaxService
 If Not fx2.IsLocalFaxPrinterInstalled Then fx2.InstallLocalFaxPrinter
End Function ' fxs_u()

Function fx1()
        Dim objFaxServer As Object 'As New FAXCOMEXLib.FaxServer
        Dim objFaxIncomingArchive As Object 'As FaxIncomingArchive
        Dim objFaxIncomingMessage As Object 'As FaxIncomingMessage

        'Error handling
        On Error GoTo Error_Handler

        'Connect to the fax server
        Set objFaxServer = CreateObject("FAXCOMEX.FaxServer")
        objFaxServer.Connect ("")

        'Get the incoming archive
        Set objFaxIncomingArchive = objFaxServer.Folders.IncomingArchive

        'Refresh the object and retrieve/display some of its properties
        Call objFaxIncomingArchive.Refresh
        MsgBox ("High quota water mark: " & objFaxIncomingArchive.HighQuotaWaterMark & _
        vbCrLf & "Low quota water mark:  " & objFaxIncomingArchive.LowQuotaWaterMark & _
        vbCrLf & "Archive folder: " & objFaxIncomingArchive.ArchiveFolder & _
        vbCrLf & "Age limit: " & objFaxIncomingArchive.AgeLimit & _
        vbCrLf & "Size high: " & objFaxIncomingArchive.SizeHigh & _
        vbCrLf & "Size low: " & objFaxIncomingArchive.SizeLow & _
        vbCrLf & "Is size quota warning on: " & objFaxIncomingArchive.SizeQuotaWarning & _
        vbCrLf & "Is archive used: " & objFaxIncomingArchive.UseArchive)

        'Set the age limit to 4 days
        objFaxIncomingArchive.AgeLimit = 4
        objFaxIncomingArchive.ArchiveFolder = "c:\eigene Dateien alt\"
        'Save the changes
        Call objFaxIncomingArchive.Save

        'Get a message by ID
        Dim MessageID As String
        Dim Answer As String
        Dim FileName As String

        Answer = InputBox("Retrieve a message by it's ID (Y/N)?")
        If Answer = "Y" Then

            MessageID = InputBox("Provide the message ID")

            'Get the job
            objFaxIncomingMessage = objFaxIncomingArchive.GetMessage(MessageID)

            'Display information about the retrieved job
            MsgBox ("Caller ID: " & objFaxIncomingMessage.CallerId & _
            vbCrLf & "CSID: " & objFaxIncomingMessage.CSID & _
            vbCrLf & "Device name: " & objFaxIncomingMessage.DeviceName & _
            vbCrLf & "Message ID: " & objFaxIncomingMessage.id & _
            vbCrLf & "Number of pages: " & objFaxIncomingMessage.Pages & _
            vbCrLf & "Number of retries: " & objFaxIncomingMessage.Retries & _
            vbCrLf & "Routing information: " & objFaxIncomingMessage.RoutingInformation & _
            vbCrLf & "Size: " & objFaxIncomingMessage.size & " bytes" & _
            vbCrLf & "Transmission start: " & objFaxIncomingMessage.TransmissionStart & _
            vbCrLf & "Transmission end: " & objFaxIncomingMessage.TransmissionEnd & _
            vbCrLf & "TSID: " & objFaxIncomingMessage.TSID)

            'Allow user to delete the message
            Answer = InputBox("Delete this message from the archive?")
            If Answer = "Y" Then Call objFaxIncomingMessage.Delete

        End If
        Exit Function

Error_Handler:
        'Implement error handling at the end of your subroutine. This implementation is for demonstration purposes
        MsgBox ("Error number: " & Hex(Err.Number) & ", " & Err.Description)

    End Function


Function fx5()
Dim objFaxServer As New FAXCOMEXLib.FaxServer ' fxscomex.dll
Dim objFaxInboundRouting 'As FaxInboundRouting
Dim collFaxInboundRoutingExtensions 'As FaxInboundRoutingExtensions
Dim objFaxInboundRoutingExtension 'As FaxInboundRoutingExtension
Dim collFaxInboundRoutingMethods 'As FaxInboundRoutingMethods
Dim collFaxInboundRoutingMethod 'As FaxInboundRoutingMethod
Dim j%
Set objFaxServer = New FAXCOMEXLib.FaxServer ' CreateObject("FAXCOMEX.FaxServer")
'Error handling
On Error GoTo Error_Handler

'Connect to the fax server
objFaxServer.Connect ""

Set collFaxInboundRoutingExtensions = objFaxServer.InboundRouting.GetExtensions
Set collFaxInboundRoutingMethods = objFaxServer.InboundRouting.GetMethods

Dim ECount As Integer
Dim MCount As Integer
'Get and display the number of routing extensions and methods on this server
ECount = collFaxInboundRoutingExtensions.Count
MCount = collFaxInboundRoutingMethods.Count
MsgBox "There are " & ECount & " routing extensions and " & _
vbCrLf & MCount & " routing methods on this server."
 
Dim n As Integer
For n = 1 To ECount
    MsgBox "Routing extension number " & n & vbCrLf & _
    vbCrLf & "Debug = " & collFaxInboundRoutingExtensions(n).Debug & _
    vbCrLf & "Name = " & collFaxInboundRoutingExtensions(n).FriendlyName & _
    vbCrLf & "Image name = " & collFaxInboundRoutingExtensions(n).ImageName & _
    vbCrLf & "Init error code = " & collFaxInboundRoutingExtensions(n).InitErrorCode & _
    vbCrLf & "Build and version = " & collFaxInboundRoutingExtensions(n).MajorBuild & "." & _
        collFaxInboundRoutingExtensions(n).MinorBuild & "." & _
        collFaxInboundRoutingExtensions(n).MajorVersion & "." & _
        collFaxInboundRoutingExtensions(n).MinorVersion & _
    vbCrLf & "Status = " & collFaxInboundRoutingExtensions(n).Status & _
    vbCrLf & "Unique name = " & collFaxInboundRoutingExtensions(n).UniqueName
    
    'Display the method GUIDs for this extension
    Dim MethodArray() As String
    MethodArray = collFaxInboundRoutingExtensions(n).Methods
    'UBound finds the size of the array
    For j = 0 To UBound(MethodArray)
        MsgBox "Routing extension number " & n & ", Method number " & j & vbCrLf & _
        MethodArray(j)
    Next

Next

Dim m As Integer
For m = 1 To MCount
    collFaxInboundRoutingMethods(m).Refresh
    MsgBox "Routing method number " & m & vbCrLf & _
    vbCrLf & "Friendly name = " & collFaxInboundRoutingMethods(m).ExtensionFriendlyName & _
    vbCrLf & "Image name = " & collFaxInboundRoutingMethods(m).ExtensionImageName & _
    vbCrLf & "Function name = " & collFaxInboundRoutingMethods(m).FunctionName & _
    vbCrLf & "Guid = " & collFaxInboundRoutingMethods(m).Guid & _
    vbCrLf & "Name = " & collFaxInboundRoutingMethods(m).Name & _
    vbCrLf & "Priority = " & collFaxInboundRoutingMethods(m).Priority
    
    'Allow change in priority
    Dim Answer As String
    Answer = InputBox("Change priority? (Y/N)")
    If Answer = "Y" Then
        Dim NewPriority As Long
        NewPriority = InputBox("Provide new priority")
        collFaxInboundRoutingMethods(m).Priority = NewPriority
        collFaxInboundRoutingMethods(m).Save
    End If
    
Next
Exit Function

Error_Handler:
    'Implement error handling at the end of your subroutine. This implementation is for demonstration purposes
    MsgBox "Error number: " & Hex(Err.Number) & ", " & Err.Description

End Function


