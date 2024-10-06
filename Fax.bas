Attribute VB_Name = "Fax"
Option Explicit
'Dim fx, fxa
Const fllNONE& = 0 'The fax server does not log events.
Const fllMIN& = 1  'The fax server logs only severe failure events, such as errors.
Const fllMED& = 2  'The fax server logs events of moderate severity, as well as severe failure events. This would include errors and warnings.
Const fllMAX& = 3  'The fax server logs all events.

Const fdrmNO_ANSWER = 0 'The device will not answer the call.
Const fdrmAUTO_ANSWER = 1 'The device will automatically answer the call.
Const fdrmMANUAL_ANSWER = 2 'The device will answer the call only if made to do so manually.
Dim FPos& ' Fehlerposition

'Private Declare Function FaxGetConfiguration Lib "winfax.dll" Alias "FaxGetConfigurationA" _
 (ByVal FaxHandle As Long, ByRef pFaxConfiguration As FAXCONFIGURATION) _
 As Long

'Private Type FAX_TIME
'    Hour As Long
'    Minute As Long
'End Type
'Private Type FAXCONFIGURATION
'    SizeOfStruct As Long
'    Retries As Long
'    RetryDelay As Long
'    Branding As Long
'    DirtyDays As Long
'    UseDeviceTsid As Long
'    ServerCp As Long
'    PauseServerQueue As Long
'    StartCheapTime As FAX_TIME
'    StopCheapTime As FAX_TIME
'    ArchiveOutgoingFaxes As Long
'    ArchiveDirectory As String
'    InboundProfile As String
'End Type
'Private hFax As Long

Function fxset%(Verz$)
 Dim fx As faxcomexlib.FaxServer      'CreateObject("FAXCOMEX.FaxServer")
 Dim fxs As faxcomlib.FaxServer       'CreateObject("faxserver.faxserver")
 Dim fxc As FAXCONTROLLib.FaxControl  'CreateObject("faxcontrol.faxcontrol")
 Dim fxps As faxcomlib.FaxPorts
 Dim fxp As faxcomlib.FaxPort
 Dim fxst As faxcomlib.FaxStatus
 Dim i%, j%
anfang:
 On Error Resume Next
 Err.Clear
 Set fxc = New FAXCONTROLLib.FaxControl
 If Err.Number <> 0 Then
  Shell (App.Path + "\nachricht.exe " + "Fehler " & Err.Number & " bei fxset: " & vbCrLf & Err.Description)
  fxset = 1
  Exit Function
 End If
 On Error GoTo fehler
 If Not fxc.IsFaxServiceInstalled Then fxc.InstallFaxService
 If Not fxc.IsLocalFaxPrinterInstalled Then fxc.InstallLocalFaxPrinter
 
 FPos = 1
 Set fxs = New faxcomlib.FaxServer
 FPos = 2
 With fxs
  .ArchiveDirectory = Verz + "\MSFax\Inbox"
  .ArchiveOutboundFaxes = -1
  .Branding = -1 ' Banner
  .DirtyDays = 3 ' wird nicht da angezeigt wo vermutet, lässt sich nicht ändern
  .Retries = 3
  .RetryDelay = 5
  .ServerCoverpage = 0
  .UseDeviceTSID = -1
 End With
 FPos = 3
 fxs.Connect ""
 FPos = 4
' GoTo anfang
 On Error Resume Next
 Set fxps = fxs.GetPorts
 If Err.Number <> 0 Then
  Shell (App.Path + "\nachricht.exe " + "Fehler " & Err.Number & " bei fxset: " & vbCrLf & Err.Description)
  fxset = 1
  Exit Function
 End If
 On Error GoTo fehler
 FPos = 5
 With fxps
  FPos = 6
  For i = 1 To .Count
   FPos = 7
   Set fxp = fxps.Item(1)
   FPos = 8
   With fxp
    FPos = 9
'   ' debug.print.CanModify
    .CSID = "GSchade 08131616381"
'   ' debug.print.Name
    ' debug.print.Priority
    .Rings = 1 ' steht bei "Geräte" sowohl innen als auch außen, zulässiger Bereich = 1-99
    .Receive = 1 ' dann geht der Haken von "Manuell" auf "Automatisch"
    .Send = -1
    .TSID = "GSchade 08131616381"
    FPos = 10
    Set fxst = fxp.GetStatus ' wohl aktuell gesendetes Fax
    FPos = 11
    Dim rms As faxcomlib.FaxRoutingMethods
    FPos = 12
    Dim rm As faxcomlib.FaxRoutingMethod
    Set rms = fxp.GetRoutingMethods
    FPos = 13
    For j = 1 To rms.Count
     FPos = 14
     Set rm = rms.Item(j)
      FPos = 15
     ' debug.printrm.DeviceId
     ' debug.printrm.DeviceName
     ' debug.printrm.Enable
     ' debug.printrm.ExtensionName
     ' debug.printrm.FriendlyName
     ' debug.printrm.FunctionName
     ' debug.printrm.Guid
     ' debug.printrm.ImageName
     ' debug.printrm.RoutingData ' hier schreibgeschützt
    Next j
   End With
  Next i
 End With
 
 FPos = 16
 Set fx = New faxcomexlib.FaxServer
 FPos = 17
 With fx
  Call .Connect(Cpt)
  FPos = 18
  With .Folders
   FPos = 19
   With .OutgoingArchive
    FPos = 20
'    .ArchiveFolder = Verz + "\MSFax\SentItems"
    .UseArchive = True
    .AgeLimit = 0 ' scheint nicht angezeigt zu werden
    .SizeQuotaWarning = False
    .HighQuotaWaterMark = -1
    .LowQuotaWaterMark = -1
    .Save
   End With
   On Error Resume Next
    With .OutgoingArchive
     .ArchiveFolder = Verz + "\MSFax\SentItems"
     FPos = 21
     .Save
    End With
   On Error GoTo fehler
   FPos = 22
   With .IncomingArchive
'    .ArchiveFolder = Verz + "\MSFax\Inbox"
    .UseArchive = True
    .AgeLimit = 0 ' scheint nicht angezeigt zu werden
    .SizeQuotaWarning = False
    .Save
   End With
   FPos = 23
   On Error Resume Next
   With .IncomingArchive
    .ArchiveFolder = Verz + "\MSFax\Inbox"
    .Save
   End With
   FPos = 24
   On Error GoTo fehler
   With .IncomingQueue
    .Blocked = False
    .Save
   End With
   FPos = 25
   With .OutgoingQueue
    .Blocked = False ' scheint nicht angezeigt zu werden
    .AgeLimit = 0 ' bei Geräte unter Bereinigen, wenn <> 0, dann angekreuzt
    .Save
   End With
  End With ' folders
  FPos = 26
  With .LoggingOptions
   With .EventLogging
    .InitEventsLevel = fllMAX
    .InboundEventsLevel = fllMAX
    .OutboundEventsLevel = fllMAX
    .GeneralEventsLevel = fllMAX
    .Save
   End With
   FPos = 27
   With .ActivityLogging
    .LogIncoming = fllMAX
    .LogOutgoing = fllMAX
    .DatabasePath = GetEnvir("allusersprofile") + "\Anwendungsdaten\Microsoft\Windows NT\MSFax\ActivityLog"
    .Save
   End With
  End With
  FPos = 28
  Dim fdev As faxcomexlib.FaxDevices
  Set fdev = fx.GetDevices
  FPos = 29
  With fdev
   For j = 0 To .Count - 1 ' wahrscheinlich müßte hier 1 stehen
    With .Item(j)
     .RingsBeforeAnswer = 1 ' ist das selbe wie oben
     .CSID = "GSchade 08131616381"
     .TSID = "GSchade 08131616381"
     .SendEnabled = -1
     .ReceiveMode = fdrmAUTO_ANSWER
     .Description = "Brother Fax von ANMELDL für MSFax"
     .Save
    End With
   Next j
  End With
  FPos = 30
  Dim fibrms As faxcomexlib.FaxInboundRoutingMethods
  Dim fibrm As faxcomexlib.FaxInboundRoutingMethod
  Set fibrms = fx.InboundRouting.GetMethods
  FPos = 31
  For i = 0 To fibrms.Count
   With fibrms.Item(i)
    ' debug.print.ExtensionFriendlyName
    ' debug.print.ExtensionImageName
    ' debug.print.FunctionName
    ' debug.print.Guid
    ' debug.print.Name
    ' debug.print.Priority
   End With
  Next i
  FPos = 32
  Dim fibres As faxcomexlib.FaxInboundRoutingExtensions
  Dim fibre As faxcomexlib.FaxInboundRoutingExtension
  Set fibres = fx.InboundRouting.GetExtensions
  For i = 0 To fibres.Count
   Set fibre = fibres(i)
   With fibre
     ' debug.print.Debug
     ' debug.print.FriendlyName
     ' debug.print.ImageName
     ' debug.print.InitErrorCode
     ' debug.print.MajorBuild
     ' debug.print.MajorVersion
     ' debug.print.Methods(0)
     ' debug.print.Methods(1)
     ' debug.print.MinorBuild
     ' debug.print.MinorVersion
     ' debug.print.Status
     ' debug.print.UniqueName
   End With
  Next i
  
 End With
 FPos = 33
 
 ' Kopie der Faxe im Patientenordner speichern
 'Dim wmireg As SWbemObjectEx
 Dim Result&, arra
 If WMIreg Is Nothing Then Set WMIreg = GetObject("winmgmts:root\default:StdRegProv")
 ' Wert eintragen
 arra = Arr(PatDokDirekt) ' c:\P
 Result = WMIreg.setbinaryvalue(HLM, "SOFTWARE\Microsoft\Fax\TAPIDevices\014BFAB1", "{92041a90-9af2-11d0-abf7-00c04fd91a4e}", arra)
 ' aktivieren
 Result = WMIreg.setbinaryvalue(HLM, "SOFTWARE\Microsoft\Fax\TAPIDevices\014BFAB1", "{aacc65ec-0091-40d6-a6f3-a2ed6057e1fa}", Array(2, 0, 0, 0))
 FPos = 34

' Call fStSpei(HLM, "SOFTWARE\Microsoft\Fax\Inbox", "Folder", EigDatDirekt + "\MSFax\Inbox")
 Call fdwSpei(HLM, "SOFTWARE\Microsoft\Fax\Inbox", "Use", 1)
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "FullName", "Gerald Schade")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "Address", "Mittermayerstraße 13" + vbCrLf + "85221 Dachau")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "City", "Dachau")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "FullName", "Gerald Schade")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "Company", "Praxis")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "Country", "Deutschland")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "FaxNumber", "08131 616381")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "HomePhone", "08131 616380")
 FPos = 35
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "mailbox", "diabetologie@dachau-mail.de")
 Call fdwSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "MonitorOnReceive", 1)
 Call fdwSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "MonitorOnSend", 1)
 Call fdwSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "NotifyIncomingCompletion", 1)
 Call fdwSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "NotifyOutgoingCompletion", 1)
 Call fdwSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "NotifyProgress", 1)
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "ZIP", "85221")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "FullName", "Gerald Schade")
' Call fDWSpei(HCU, "SOFTWARE\Microsoft\Fax", "Dirty Days", 0) ' zu ändern über outgoing folder
' Call fStSpei(HCU, "SYSTEM\CurrentControlSet\Control\Print\Printers\Fax", "Location", cpt) ' zu ändern was weiß ich wo
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fxset/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' fxset
Function Arr(Pfad$) ' Baut nach jedem Buchstaben eine Lücke ein, so wie es offenbar die Routing-Funktion will
 Dim i%, j%, arra%()
 On Error GoTo fehler
 ReDim arra(2 * Len(Pfad))
 j = 0
 For i = 0 To Len(Pfad) - 1
  arra(j) = Asc(Mid(Pfad, i + 1, 1))
  arra(j + 1) = 0
  j = j + 2
 Next i
 Arr = arra
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Arr/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Arr

Sub FaxSend()
Dim FaxServer As New faxcomlib.FaxServer
Dim FaxDoc As New faxcomlib.FaxDoc
Dim FaxTiff As New faxcomlib.FaxTiff
Dim strFaxJob As faxcomlib.FaxJobs
Dim strFaxStatus As faxcomlib.FaxJob
Dim DateiName$
Dim strFaxTiff As faxcomlib.FaxTiff
Dim strJobID&
On Error GoTo fehler
 DateiName = "u:\test1.txt"
 Err.Clear
 FaxServer.Connect ("ANMELDL")

Set FaxDoc = FaxServer.CreateDocument(DateiName)
   
    FaxDoc.BillingCode = "Rechnungsnummer 381"
    FaxDoc.CoverpageName = ""
    FaxDoc.CoverpageNote = "Note von Gerald"
    FaxDoc.CoverpageSubject = "Thema der Coverpage"
    FaxDoc.DiscountSend = 0
    FaxDoc.DisplayName = "Fax von mir"
    FaxDoc.EmailAddress = "diabetologie@dachau-mail.de"
    FaxDoc.FaxNumber = "08131 619713"
    FaxDoc.RecipientAddress = "Teerstraße 15"
    FaxDoc.RecipientCity = "Odelzhausen"
    FaxDoc.RecipientCompany = "Hausarztpraxis"
    FaxDoc.RecipientCountry = "D"
    FaxDoc.RecipientDepartment = "Archiv"
    FaxDoc.RecipientHomePhone = "666414"
    FaxDoc.RecipientName = "Gerald Schade auf ANMELDL"
    FaxDoc.RecipientOffice = "Büro"
    FaxDoc.RecipientOfficePhone = "616380"
    FaxDoc.RecipientState = "Bayern"
    FaxDoc.RecipientTitle = "DrHC"
    FaxDoc.RecipientZip = "85221"
    FaxDoc.SendCoverpage = 0
    FaxDoc.SenderAddress = "Holzweg 1"
    FaxDoc.SenderCompany = "Diabetologische Schwerpunktpraxis"
    FaxDoc.SenderDepartment = "Schreibbüro"
    FaxDoc.SenderFax = "08131 616381"
    FaxDoc.SenderHomePhone = "9037022"
    FaxDoc.SenderName = "Gerald Schade"
    FaxDoc.SenderOffice = "Praxis"
    FaxDoc.SenderOfficePhone = "76373"
    FaxDoc.SenderTitle = "Prof.HC"
    FaxDoc.ServerCoverpage = 1
    strJobID = FaxDoc.Send
    
'    MsgBox FaxServer.ArchiveDirectory
  
Set strFaxJob = FaxServer.GetJobs()
Set strFaxStatus = strFaxJob.Item(1)
    
On Error Resume Next

Set FaxServer = Nothing
Set FaxDoc = Nothing
Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in FaxSend/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' FaxSend

Function fxp()
Dim myptr As Printer
On Error GoTo fehler
For Each myptr In Printers
   If myptr.DeviceName = "Fax" Then
      Set Printer = myptr
      Exit For
   End If
Next
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fxp/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select

End Function ' fxp

