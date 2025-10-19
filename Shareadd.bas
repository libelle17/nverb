Attribute VB_Name = "ShareaddMod"
Private Const NERR_SUCCESS As Long = 0&

'share types
Private Const STYPE_ALL As Long = -1  'note: my const
Private Const STYPE_DISKTREE As Long = 0
Private Const STYPE_PRINTQ As Long = 1
Private Const STYPE_DEVICE As Long = 2
Private Const STYPE_IPC As Long = 3
Private Const STYPE_SPECIAL As Long = &H80000000

'permissions
Private Const ACCESS_READ As Long = &H1
Private Const ACCESS_WRITE As Long = &H2
Private Const ACCESS_CREATE As Long = &H4
Private Const ACCESS_EXEC As Long = &H8
Private Const ACCESS_DELETE As Long = &H10
Private Const ACCESS_ATRIB As Long = &H20
Private Const ACCESS_PERM As Long = &H40
Private Const ACCESS_ALL As Long = ACCESS_READ Or _
                                   ACCESS_WRITE Or _
                                   ACCESS_CREATE Or _
                                   ACCESS_EXEC Or _
                                   ACCESS_DELETE Or _
                                   ACCESS_ATRIB Or _
                                   ACCESS_PERM

Private Type SHARE_INFO_2
  shi2_netname As Long
  shi2_type As Long
  shi2_remark As Long
  shi2_permissions As Long
  shi2_max_uses As Long
  shi2_current_uses  As Long
  shi2_path As Long
  shi2_passwd As Long
End Type
  
Declare Function NetShareDelNT& Lib "netapi32.dll" Alias "NetShareDel" (ByVal servername&, ByVal netname$, ByVal Reserved&)
Declare Function NetShareAdd& Lib "netapi32" (ByVal servername&, ByVal level&, buf As Any, parmerr&)


Function Shareadd&(ByVal Server$, ByVal FreigabePfad$, ByVal FreigabeName$, Optional ByVal Remark$, Optional ByVal FreigabePW$, Optional ByVal ReadOnly%)
' net share Multimedia=m:\Multimedia /remark:"Multimedia"
' cacls m:\Multimedia /E /P florian:F
  Dim dwServer   As Long
  Dim dwNetname  As Long
  Dim dwPath     As Long
  Dim dwRemark   As Long
  Dim dwPw       As Long
  Dim parmerr    As Long
  Dim si2        As SHARE_INFO_2
  Dim gabsschon%, fgpfad$
  If Not FSO.FolderExists(FreigabePfad) Or Left(FreigabePfad, 1) = "\" Then Exit Function
  If Right$(FreigabePfad, 1) = "\" Then
   If Right$(FreigabePfad, 2) = ":\" Then
    fgpfad = FreigabePfad
   Else
    fgpfad = Left(FreigabePfad, Len(FreigabePfad) - 1)
   End If
  Else
   If Right$(FreigabePfad, 1) = ":" Then
    fgpfad = FreigabePfad & "\"
   Else
    fgpfad = FreigabePfad
   End If
  End If
  On Error GoTo fehler
  gabsschon = FSO.FolderExists(Server & "\" & FreigabeName)
  
  If WV < win_vista Then
'  If Server = "" Then Server = "\\" + Cpt
  ' Setze Pointer zu Server Freigabe Pfad und Name
'  dwServer = StrPtr(IIf(Server = "", "\\" + Cpt, Server))
  dwServer = StrPtr(Server)
  dwNetname = StrPtr(FreigabeName)
  dwPath = StrPtr(fgpfad)
   
  ' Wenn Remark und Passwort gesetzt sind,
  ' setze Pointer zur Auswahl
  If Len(Remark) > 0 Then
    dwRemark = StrPtr(Remark)
  End If
   
  If Len(FreigabePW) > 0 Then
    dwPw = StrPtr(FreigabePW)
  End If
      
  ' SHARE_INFO_2 Struktur erstellen
  With si2
    .shi2_netname = dwNetname
    .shi2_path = dwPath
    .shi2_remark = dwRemark
    .shi2_type = STYPE_DISKTREE
    .shi2_permissions = ACCESS_ALL
    If ReadOnly Then .shi2_permissions = (ACCESS_PERM Or ACCESS_READ)
    .shi2_max_uses = -1
    .shi2_passwd = dwPw
  End With
  If gabsschon Then
   Dim nerr&
   nerr = NetShareDelNT(0&, StrConv(FreigabeName, vbUnicode), 0&)
  End If
                          
  ' Freigabe hinzufügen
   Shareadd = NetShareAdd(dwServer, 2, si2, parmerr)
 Else ' win_ver < vista
  If Not gabsschon Then
   If Right(FreigabePfad, 1) = "\" Then FreigabePfad = Left(FreigabePfad, Len(FreigabePfad) - 1)
'   erg = Not Shell(doalsad & acceu & AdminGes & " cmd /c net share " & FreigabeName & "=" & Chr$(34) & FreigabePfad & Chr$(34) & " /unlimited /remark:" & Chr$(34) & Remark & Chr$(34))
'   erg = Not SuSh("cmd /c net share " & FreigabeName & "=" & Chr$(34) & FreigabePfad & Chr$(34) & " /unlimited /remark:" & Chr$(34) & Remark & Chr$(34), 2)
   erg = rufauf("cmd", "/c net share " & FreigabeName & "=""" & FreigabePfad & """ /unlimited /remark:""" & Remark & """", 2, , , 0)
  End If
 End If
 If parmerr <> 0 Then
'   Call Shell(App.Path & "\nachricht.exe " & "Fehler " & parmerr & " bei NetShareAdd" & vbCrLf & "Server: " & Server & vbCrLf & "Freigabepfad: " & fgpfad & vbCrLf & "Freigabename: " & FreigabeName)
'    SuSh App.Path & "\nachricht.exe " & "Fehler " & parmerr & " bei NetShareAdd" & vbCrLf & "Server: " & Server & vbCrLf & "Freigabepfad: " & fgpfad & vbCrLf & "Freigabename: " & FreigabeName, 0, , 0, 1
    rufauf App.Path & "\nachricht.exe", "Fehler " & parmerr & " bei NetShareAdd" & vbCrLf & "Server: " & Server & vbCrLf & "Freigabepfad: " & fgpfad & vbCrLf & "Freigabename: " & FreigabeName, , , 0, 1, 0
   Else
'   Stop
   End If
   If gabsschon Then
   Else
'   Call Shell(App.Path & "\berechtigungen.exe " & fgpfad)
'   Call SuSh(App.Path & "\berechtigungen.exe " & fgpfad, 2, , 0, 1)
    rufauf App.Path & "\berechtigungen_direkt.exe", fgpfad, , , 0, 1
'   Dim erg%
'   erg = MsgBox("Sollen die Berechtigungen für " & FreigabePfad & " gesetzt werden?", vbYesNo)
'   If erg = vbYes Then
'    Call Shell("cacls " & FreigabePfad & " /T /E /P Jeder:C")
'   End If
 End If
 Exit Function
fehler:
Select Case MsgBox("Fehler in ShareAdd: " & "Server:" & Server & vbCrLf & "Freigabename:" & FreigabeName & vbCrLf & "Freigabepfad: " & fgpfad & vbCrLf & "Remark:" & Remark & vbCrLf & "FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ShareAdd/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ShareAdd

