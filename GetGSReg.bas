Attribute VB_Name = "GetGSReg"
Option Explicit

Function fgetgsreg(NeuServ$)
' Const NeuServ = "MITTE"
Const ANMR$ = "ANMELDR1" '"ANMELDRNEU"
 Dim obR As Object, rP$, aVN0, aVN1, i%, j%, k%, erg, phkResult&
 Dim aServ
 Dim dlen&, data$
 dlen = 2048
 Const strComputer = "."
 On Error GoTo fehler
 rP = "Software\GSProducts"
 If 1 = 0 Then
 On Error Resume Next
 Set obR = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 If Err.Number <> 0 Then
  On Error GoTo fehler
  Err.Clear
  obR.Enumkey HCU, rP, aVN0
  For i = 0 To UBound(aVN0)
   On Error Resume Next
   aServ = Null
   obR.GetStringValue HCU, rP & "\" & aVN0(i) & "\DBVerb", ANMR, aServ
   If Not IsNull(aServ) And aServ <> NeuServ Then
    obR.setStringValue HCU, rP & "\" & aVN0(i) & "\DBVerb", ANMR, NeuServ
   End If
   On Error GoTo fehler
   obR.Enumkey HCU, rP & "\" & aVN0(i) & "\DBVerb", aVN1
   If Not IsNull(aVN1) Then
    For j = 0 To UBound(aVN1)
     On Error Resume Next
     aServ = Null
     obR.GetStringValue HCU, rP & "\" & aVN0(i) & "\DBVerb" & "\" & aVN1(j), ANMR, aServ
     If Not IsNull(aServ) And aServ <> NeuServ Then
      obR.setStringValue HCU, rP & "\" & aVN0(i) & "\DBVerb" & "\" & aVN1(j), ANMR, NeuServ
     End If
    Next j
   End If
  Next i
 Else
  On Error GoTo fehler
  Err.Clear
 End If
Else ' 1 = 0
  Dim v1$(), v2$()
  Call regEnumSub(HCU, rP, v1)
  For i = 1 To UBound(v1)
   Debug.Print v1(i)
   aServ = getReg(HCU, rP & "\" & v1(i) & "\DBVerb", ANMR)
   If Not IsNull(aServ) And Not IsEmpty(aServ) And aServ <> NeuServ Then
    Call fStSpei(HCU, rP & "\" & v1(i) & "\DBVerb", ANMR, NeuServ)
   End If
   Call regEnumSub(HCU, rP & "\" & v1(i) & "\DBVerb" & "\", v2)
   For j = 1 To UBound(v2)
    aServ = getReg(HCU, rP & "\" & v1(i) & "\DBVerb" & "\" & v2(j), ANMR)
    If Not IsNull(aServ) And Not IsEmpty(aServ) And aServ <> NeuServ Then
     Call fStSpei(HCU, rP & "\" & v1(i) & "\DBVerb" & "\" & v2(j), ANMR, NeuServ)
    End If
   Next j
  Next i
  
 End If
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fgetgsreg/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' fgetgsreg


Function fragReg(Zuord&, Schlüssel$, Wert$) As Variant
Dim erg& ' Ergebnisse von RegOpenKeyEx und RegQueryValueEx
Dim hKey&
Dim lpSubKey$
Dim ulOptions&
Dim samDesired&
Dim phkResult&
Dim lpValueName$
Dim lpReserved&
Dim lpType&
Dim lpData As String * 100
Dim lpcbData&
Dim i%
On Error GoTo fehler
Select Case Zuord
 Case 0, &H80000000
    hKey = &H80000000 ' HKEY_CLASSES_ROOT ' steht in winreg.h
 Case 1, &H80000001
    hKey = &H80000001 ' HKEY_CURRENT_USER ' steht in winreg.h
 Case 2, &H80000002
    hKey = &H80000002 ' HKEY_LOCAL_MACHINE ' steht in winreg.h
 Case 3, &H80000003
    hKey = &H80000003 ' HKEY_USERS ' steht in winreg.h
 Case 4, &H80000004
    hKey = &H80000004 ' HKEY_PERFORMANCE_DATA' steht in winreg.h
 Case 5, &H80000005
    hKey = &H80000005 ' HKEY_CURRENT_CONFIG' steht in winreg.h
 Case 6, &H80000006
    hKey = &H80000006 ' HKEY_DYN_DATA' steht in winreg.h
 Case Else
    hKey = Zuord
End Select
lpSubKey = Schlüssel
ulOptions = 0
samDesired = &H20119 ' ' steht in winnt.h -> KEY_READ (alles zusammenzählen)
'Const KEY_READ& = &H20019
'Const KEY_WOW64_64KEY& = &H100
erg = RegOpenKeyEx(hKey, lpSubKey, ulOptions, samDesired, phkResult)
If erg = 0 Then ' Debug.Print "Success=", erg, "phkResult=", phkResult
 lpValueName = Wert
 lpReserved = 0
 lpType = 1
 lpcbData = 100
 erg = RegQueryValueEx(phkResult, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
 If lpType <> 1 Then
  erg = RegQueryValueEx(phkResult, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
 End If
 fragReg = RegTrim$(lpData, lpcbData)
End If 'erg = 0
Exit Function
fehler:
ErrNumber = Err.Number
ErrDescription = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) & vbCrLf & "Description: " & ErrDescription & vbCrLf & "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fragReg/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetReg

