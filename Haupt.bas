Attribute VB_Name = "Haupt"
' unter "Projekt -> Verweise" muß "Edanmo's Task Scheduler Class v1.10" auf "tskschd.dll" zeigen
Option Explicit
Public IrfanVerz$, IrfanExe$, IrfanPfad$, IrfanErg%
Public ErrNumber&, ErrDescription$, ErrSource$, ErrLastDllError$

Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long
Declare Function GetAtomName Lib "kernel32.dll" Alias "GetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As cRECT) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Type cRECT
  Left As Long
  top As Long
  Right As Long
  bottom As Long
End Type
Private Declare Function GetClassInfoEx Lib "user32.dll" _
  Alias "GetClassInfoExA" ( _
  ByVal hinst As Long, _
  ByVal lpszClass As String, _
  lpwcx As WNDCLASSEX) As Long
Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" ( _
  ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" _
  Alias "GetClassNameA" ( _
  ByVal hwnd As Long, _
  ByVal lpClassName As String, _
  ByVal nMaxCount As Long) As Long
Private Type WNDCLASSEX
  cbSize As Long
  style As Long
  lpfnWndProc As Long
  cbClsExtra As Long
  cbWndExtra As Long
  hInstance As Long
  hIcon As Long
  hCursor As Long
  hbrBackground As Long
  lpszMenuName As String
  lpszClassName As String
  hIconSm As Long
End Type
' GetWindow wCmd-Konstanten
Private Const GW_HWNDFIRST = 0 ' Ermittelt das erste Fenster aus der Z-Order,
' in dem sich das angegebene Fenster befindet
Private Const GW_HWNDLAST = 1 ' Ermittelt das letzte Fenster aus der Z-Order,
' in dem sich das angegebene Fenster befindet
Private Const GW_HWNDNEXT = 2 ' Ermittelt das nächste Fenster aus der Z-Order,
' in dem sich das angegebene Fenster befindet
Private Const GW_HWNDPREV = 3 ' Ermittelt das vorherige Fenster aus der Z-Order,
' in dem sich das angegebene Fenster befindet
Private Const GW_OWNER = 4 ' Ermittelt das Fensterhandle des Fenster, welches
' dem angegebenen übergeordnet ist
Private Const GW_CHILD = 5 ' Ermittelt das Fensterhandle des Kindfensters,
' welches sich im Vordergrund befindet und / oder den Focus besitzt
 
' eine der GetWIndowLong nIndex-Konstanten
Private Const GWL_HINSTANCE = (-6)

Const WM_SETREDRAW = &HB
Const RDW_INVALIDATE = &H1
Const RDW_INTERNALPAINT = &H2
Const RDW_ERASE = &H4
Const RDW_VALIDATE = &H8
Const RDW_NOINTERNALPAINT = &H10
Const RDW_NOERASE = &H20
Const RDW_NOCHILDREN = &H40
Const RDW_ALLCHILDREN = &H80
Const RDW_UPDATENOW = &H100
Const RDW_ERASENOW = &H200
Const RDW_FRAME = &H400
Const RDW_NOFRAME = &H800

Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As cRECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'Public WMIreg As SWbemObjectEx ' %windir%\system32\wbem\wbemdisp.tlb
Public FPos& ' Fehlerposition
Public Const vNS$ = vbNullString
' Public uVerz$, vVerz$, pVerz$, pDatenb$, plzVz$ ' "P:\datenbanken" ' 30.12.22 auskommentiert, da auch in ComputerTools
Public pDatenb$ ' "P:\datenbanken"

Const obSchottdorf% = 0
Const obStaber% = True
Public IsIDE%
Public FNr&
Public RegOrt& ' HCU oder HLM, je nach Windows-Version
Public oboffenlassen%
'Dim oEnvSystem As New System
Dim Fl As File
' Dim FI As New FürIcon
' Public Declare Function FaxGetConfiguration Lib "winfax.dll" (ByVal FaxHandle As Long, ByRef FaxConfig As FAX_CONFIGURATION) As Long
  Const ERROR_MORE_DATA = 234
  Const ERROR_SUCCESS As Long = 0
  
    Declare Function EnableTheming Lib "UxTheme.dll" (ByVal fEnable As Boolean) As Integer

    Declare Function IsThemeActive Lib "UxTheme.dll" () As Integer

    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Integer, ByVal uiParam As Integer, ByRef pvParam As Integer, ByVal fWinIni As Integer) As Integer
    ' Konstanten für uAction
    Const SPI_GETACCESSTIMEOUT As Integer = 60
    ' empfängt die Maximale Zeit die der Benutzer keine Eingaben macht, wobei dann
    ' bestimmte Funktionen deaktiviert werden. "uiParam" muss die Größe der Struktur
    ' ACCESSTIMEOUT sein und "pvParam" muss eine ACCESSTIMEOUT-Struktur übergeben werden,
    ' die gefüllt wird

    Const SPI_GETANIMATION As Integer = 72
    ' empfängt Informationen über Effekte, die bei bestimmten Benutzereingaben aufgerufen werden.
    '  "uiParam" muss `0` sein und "pvParam" muss eine ANIMATIONINFO-Strukur
    ' erhalten, die gefüllt werden soll.

    Const SPI_GETBEEP As Integer = 1
    ' ermittelt, ob der Beeper aktviert (ungleich 0) oder deaktiviert ist (0). "uiParam"
    ' muss eine `0` übergeben werden und "pvParam" erwartet eine Long-Variable, die mit
    ' dem ermittelten Wert gefüllt wird.

    Const SPI_GETBORDER As Integer = 5
    ' ermittelt den Faktorwert, die ein in der Größe veränderbarer Fensterrahmen dick ist.
    ' "uiParam" muss eine `0` übergeben werden und "pvParam" erwartet eine Long-Variable,
    ' die mit dem Wert gefüllt wird.

    Const SPI_GETDEFAULTINPUTLANG As Integer = 89
    ' (nur Win 9x) ermittelt das Handle des Keyboard-Layouts, das für die Standard-Systemsprache
    ' benutzt wird. "uiParam" muss `0` sein und "pvParam" erwartet eine Long-Variable,
    ' die das Handle empfängt

    Const SPI_GETDRAGFULLWINDOWS As Integer = 38
    ' (nur Win 9x) ermittelt, ob der Fensterinhalt beim Verschieben eines Fensters
    ' angezeigt (ungleich 0) oder nicht angezeigt (0) werden soll. "uiParam" muss `0`
    ' sein und "pvParam" erwartet eine Long-Variale, die mit dem Wert gefüllt wird

    Const SPI_GETFASTTASKSWITCH As Integer = 35
    ' ermittelt ob das Wechseln zwischen den Fenstern mit ALT+TAB möglich (ungleich 0)
    ' ist oder ob dies deaktiviert (0) ist. "uiParam" erwartet `0` und "pvParam"
    ' erwartet eine Long-Variable, die mit dem entsprechenden Wert gefüllt wird.

    Const SPI_GETFILTERKEYS As Integer = 50
    ' ermittelt Informationen über die Filterkeys die z.B. die Wiederholungsrate der
    ' Buchstaben beim Drücken einer Taste der Tastatur festlegen. "uiParam" erwartet
    ' die Größe der FILTERKEYS-Struktur und "pvParam" erwartet eine FILTERKEYS-Struktur,
    ' die mit den Informationen gefüllt wird.

    Const SPI_GETFONTSMOOTHING As Integer = 74
    ' ermittelt, ob Fonts windowsweit weich (ungleich 0) oder normal (0) gezeichnet
    ' werden. "uiParam" muss `0` sein und "pvParam" erwartet eine Long-Variable, die das
    ' Ergebnis enthält.

    Const SPI_GETGRIDGRANULARITY As Integer = 18
    ' ermittelt die aktuelle Desktop-Gitternetzkörnung. "uiParam" muss `0` sein und
    ' "pvParam" erwartet eine Long-Variable, die den ermittelten Wert enthält.

    Const SPI_GETHIGHCONTRAST As Integer = 66
    ' (nur Win 9x) ermittelt die HighContrast-Einstellungen die für die bessere
    ' Sichtbarkeit der Fenster verantwortlich ist. "uiParam" muss `0` sein und "pvParam"
    ' erwartet eine HIGHCONTRAST-Struktur, die mit den Daten gefüllt wird.

    Const SPI_GETICONMETRICS As Integer = 45
    ' (nur Win 9x) ermittelt Informationen wie Icons angezeigt werden. "uiParam" muss `0` sein
    ' und "pvParam" erwartet eine ICONMETRICS Struktur, die mit den Daten gefüllt wird.

    Const SPI_GETICONTITLELOGFONT As Integer = 31
    ' ermittelt Informationen über die Schriftart, die für Iconbeschriftung benutzt wird.
    ' "uiParam" erwartet die Größe der LOGFONT-Struktur und "pvParam" erwartet eine
    ' LOGFONT-Struktur, die mit den Font-Informationen gefüllt wird.

    Const SPI_GETICONTITLEWRAP As Integer = 25
    ' ermittelt, ob zu lange Titel der Icons umgebrochen (ungleich 0) werden oder
    ' vollständig (0) in einer Zeile angezeigt werden. "uiParam" muss `0` sein und
    ' "pvParam" erwartet eine Long-Variable, die den Wert enthält.

    Const SPI_GETKEYBOARDDELAY As Integer = 22
    ' ermittelt das Zeitlimit, nach dem eine gedrückte Taste der Tastatur wiederholt
    ' wird. "uiParam" muss `0` sein und "pvParam" erwartet eine Long-Variable, die mit
    ' einem Wert zwischen `0` und `3` gefüllt wird.

    Const SPI_GETKEYBOARDPREF As Integer = 68
    ' (nur Win 9x) ermittelt, ob sich der Benutzer auf die Tastatur verlässt anstatt auf
    ' die Maus. "uiParam" muss `0` sein und "pvParam" empfängt `0`, wenn der Benutzer die
    ' Tastatur benutzt und `ungleich 0` wenn er die Maus benutzt.

    Const SPI_GETKEYBOARDSPEED As Integer = 10
    ' ermittelt die Wiederholungsrate einer gedrückten Tastaturtaste. "uiParam" muss `0`
    ' sein und "pvParam" erwartet eine Long-Variable, die mit einem Wert zwischen `0` und `31`
    ' gefüllt wird.

    Const SPI_GETLOWPOWERACTIVE As Integer = 83
    ' (nur Win 9x)  ermittelt, ob sich das System in einem Standbybetrieb (ungleich 0)
    ' befindet oder nicht (0). "uiParam" muss `0` sein und "pvParam" erwartet eine
    ' Long-Variable, die das Ergebnis empfängt.

    Const SPI_GETLOWPOWERTIMEOUT As Integer = 79
    ' (nur Win 9x) ermittelt die Zeit in Sekunden, nach dem das System den Standbybetrieb
    ' aktiviert wenn keine Benutzereingabe erfolgte. "uiParam" muss `0` sein
    ' und "pvParam" erwartet eine Long-Variable, die den Wert enthält.

    Const SPI_GETMENUDROPALIGNMENT As Integer = 27
    '  ermittelt, ob Popupmenüs links (ungleich  0) oder rechts (0) ausgerichtet werden.
    ' "uiParam" muss `0` sein und "pvParam" erwartet eine Long-Variable, die den Wert enthält.

    Const SPI_GETMINIMIZEDMETRICS As Integer = 43
    ' (nur Win 9x) ermittelt die Ausrichtung, mit der minimierte Fenster angezeigt werden.
    ' "uiParam" erwartet die Größe der MINIMIZEDMETRICS-Struktur und "pvParam" erwartet
    ' eine MINIMIZEDMETRIC-Struktur, die mit den Daten gefüllt wird.

    Const SPI_GETMOUSE As Integer = 3
    ' ermittelt die Einstellungen der Maus X- & Y-Achse und der Mausgeschwindigkeit.
    ' "uiParam" muss `0` sein und "pvParam" erwartet ein 3 Felder großes Long-Array,
    ' das die X-Achsen, Y-Achsen und Mausspeed-Einstellungen enthält.

    Const SPI_GETMOUSEKEYS As Integer = 54
    ' ermittelt Informationen über die MouseKeys, die es ermöglichen die Maus per NumPad
    ' zu steuern. "uiParam" erwartet die Größe der MOUSEKEYS-Struktur und "pvParam"
    ' erwartet die MOUSEKEYS Struktur die mit den Informationen gefüllt wird.

    Const SPI_GETMOUSETRAILS As Integer = 94
    ' (nur Win 9x) ermittelt, ob eine Mausspur vorhanden (größer als 1) ist.

    Const SPI_GETNONCLIENTMETRICS As Integer = 41
    ' (Nur Win 9x) ermittelt die Eigenschaften eines MDI-Fensterbereichs. "uiParam"
    ' muss `0` sein und "pvParam" erwartet eine NONCLIENTMETRICS- Struktur, die gefüllt wird.

    Const SPI_GETPOWEROFFACTIVE As Integer = 84
    ' (nur Win 9x) ermittelt, ob das System den PowerOff-Modus erreicht hat, nachdem der
    ' Benutzer einen bestimmten Zeitraum keine Eingaben gemacht hat. "uiParam" muss `0`
    ' sein und "pvParam" erwartet eine Long-Variable, die mit `0` gefüllt wird, wenn der
    ' PowerOff-Modus nicht aktiv ist oder andernfalls `ungleich 0`

    Const SPI_GETPOWEROFFTIMEOUT As Integer = 80
    ' (nur Win 9x) ermittelt den Zeitraum, den das System auf eine Benutzereingabe wartet
    ' bis der PowerOff-Modus gestartet wird. "uiParam" muss `0` sein und "pvParam"
    ' erwartet eine Long-Variable, die mit der Anzahl der Sekunden gefüllt wird.

    Const SPI_GETSCREENREADER As Integer = 70
    ' (nur Win 9x) ermittelt, ob ein ScreenReader aktiviert (ungleich 0) ist oder nicht (0),
    ' um eine Anwendung mehr textbasiernd zu gestalten. "uiParam" muss `0` sein und
    ' "pvParam" erwartet eine Long-Variable, die mit dem Wert gefüllt wird.

    Const SPI_GETSCREENSAVEACTIVE As Integer = 16
    ' ermittelt, ob der Bildschirmschoner momentan ausgeführt (ungleich 0) wird oder
    ' nicht (0). "uiParam" muss `0` sein und "pvParam" erwartet eine Long-Variable, die
    ' mit dem ermitteltem Wert gefüllt wird.

    Const SPI_GETSCREENSAVETIMEOUT As Integer = 14
    ' ermittelt die Anzahl an Sekunden, nach denen der Bildschirmschoner gestartet wird.
    ' "uiParam" muss `0` sein und "pvParam" erwartet eine Long-Variable, die mit der
    ' Anzahl der Sekunden gefüllt wird.

    Const SPI_GETSERIALKEYS As Integer = 62
    ' ermittelt Informationen über die SerialKeys, die wiederum die Ports steuern um z.B.
    ' Tastatur oder Maus anzusteuern. "uiParam" muss `0` sein und "pvParam" erwartet
    ' eine SERIALKEYS Struktur, die mit den Informationen gefüllt wird.

    Const SPI_GETSHOWSOUNDS As Integer = 56
    ' ermittelt, ob der Benutzer vorzugsweise zusätzlich (ungleich 0) visuelle oder nur
    ' (0) akustische Signale verwendet. "uiParam" muss `0` sein und "pvParam" erwartet
    ' eine Long-Variable, die mit dem Wert gefüllt wird.

    Const SPI_GETSOUNDSENTRY As Integer = 64
    ' ermittelt die Eigenschaften der SOUNDSENTRY, die akustische Effekte mit visuellen
    ' untermalen. "uiParam" erwartet die Größe der SOUNDSENTRY-Struktur und "pvParam"
    ' erwartet die SOUNDSENTRY-Struktur, die mit den Informationen gefüllt wird.

    Const SPI_GETSTICKYKEYS As Integer = 58
    ' ermittelt Informationen über STICKKEYS, die das Drücken von 2 Tastaturtasten
    ' gleichzeitig dadurch vereinfachen, dass auch zeitliche Differenzen zwischen dem
    ' Drücken der einen und Loslassen der anderen Taste akzeptiert werden. "uiParam"
    ' erwartet die Größe der STICKKEYS-Struktur und "pvParam" erwartet die STICKKEYS-
    ' Struktur, die mit den Informationen gefüllt wird.

    Const SPI_GETTOGGLEKEYS As Integer = 52
    ' ermittelt Informationen über TOGGELKEYS, die beim Drücken der Num-, Feststell- und
    ' Rollen-Taste einen Sound abspielen. "uiParam" erwartet die Größe der TOGGELKEYS-
    ' Struktur und "pvParam" erwartet die TOGGELKEYS-Struktur, die mit den Informationen
    ' gefüllt wird.

    Const SPI_GETWINDOWSEXTENSION As Integer = 92
    ' (nur Win 9x) ermittelt ob Windows Extensions installiert ist oder nicht. Windows Extensions
    ' ist ein Teil von Windows PLUS. "uiParam" muss `1` sein und "pvParam" muss `0` sein,
    ' die Funktion gibt `1` zurück, wenn die Extensions installiert sind, andernfalls `0`.

    Const SPI_GETWORKAREA As Integer = 48
    ' (nur Win 9x) ermittelt den Arbeitsbereich des Desktops abzüglich der Taskbar.
    ' "uiParam" muss `0` sein und "pvParam" erwartet eine Rect-Struktur, die mit den
    ' Koordinaten gefüllt wird.

    Const SPI_ICONHORIZONTALSPACING As Integer = 13
    ' setzt die neue Breite einer Icon-Zelle. "uiParam" erwartet den neuen Wert und
    ' "pvParam" muss `0` sein.

    Const SPI_ICONVERTICALSPACING As Integer = 24
    ' setzt die neue Höhe einer Icon-Zelle. "uiParam" erwartet den neuen Wert und
    ' "pvParam" muss `0` sein.

    Const SPI_LANGDRIVER As Integer = 12
    ' (nur Win 9x) ermittelt den Dateinamen des Sprachtreibers. "uiParam" muss `0` sein
    ' und "pvParam" erwartet einen String-Puffer mit vorinitialisierten Leerzeichen, um
    ' den String zu erhalten

    Const SPI_SCREENSAVERRUNNING As Integer = 97
    ' (nur Win 9x) deaktiviert die Tastenkombinationen STRG + ALT + ENTF und ALT + TAB.
    ' "uiParam" erwartet einen BOOLESCHEN Wert, um die Tastenkombinationen zu aktivieren
    ' (True) oder zu deaktivieren (False), und "pvParam" erwartet den BOOLESCHEN Wert ,der
    ' vorher gesetzt war.

    Const SPI_SETACCESSTIMEOUT As Integer = 61
    ' setzt Informationen zu den ACCESSEDTIMEOUT-Eigenschaften. "uiParam" erwartet die Größe der
    ' ACCESSEDTIMEOUT-Struktur, die in "pvParam" übergeben werden muss.

    Const SPI_SETANIMATION As Integer = 73
    ' (nur Win 9x) setzt Informationen über Effekte, wenn ein Fenster verschoben,
    ' minimiert oder maximiert wird. "uiParam" muss `0` sein und "pvParam" erwartet eine
    ' gefüllte ANIMATIONINFO-Struktur.

    Const SPI_SETBEEP As Integer = 2
    ' schaltet den Systembeeper an (ungleich 0) oder aus (0). "uiParam" erwartet den
    ' Wert, um den Status des Beepers zu verändern und "pvParam" muss `0` sein.

    Const SPI_SETBORDER As Integer = 6
    ' setzt den Multiplizierungs-Faktor eines in der Größe veränderbaren Rahmens.
    ' "uiParam" erwartet den neuen Wert und "pvParam" muss `0` sein.

    Const SPI_SETCURSORS As Integer = 87
    ' zeichnet den Systemcursor neu. "uiParam" und "pvParam" müssen `0` sein.

    Const SPI_SETDEFAULTINPUTLANG As Integer = 90
    ' (nur Win 9x) setzt das Layout das für die Standard-Systemsprache verwendet werden
    ' soll.  "uiParam" erwartet das Handle eines Tastatur-Layouts und "pvParam" muss `0` sein.

    Const SPI_SETDESKPATTERN As Integer = 21
    ' lädt das Desktophintergrundbild erneut mit dem Anzeigeformat, das in der Win.ini
    ' unter "Desktop" bei "Pattern" eingetragen ist. "uiParam" und "pvParam" müssen `0` sein.

    Const SPI_SETDESKWALLPAPER As Integer = 20
    ' setzt das Desktophintergrund-Bitmap. "uiParam" muss `0` sein und bei "pvParam" muss
    ' der Dateipfad des neuen Hintergrundbilds als String übergeben werden."

    Const SPI_SETDOUBLECLICKTIME As Integer = 32
    ' setzt die Zeit in Millisekunden, die maximal vergehen darf ,damit Windows einen
    ' Doppelklick erkennt. "uiParam" erwartet den neuen Wert und "pvParam" muss `0` sein.

    Const SPI_SETDOUBLECLKHEIGHT As Integer = 30
    ' setzt die maximale Höhe, die sich der Mauszeiger bewegen darf, damit Windows einen
    ' Doppelklick erkennt. "uiParam" erwartet den neuen Wert und "pvParam" muss `0` sein.

    Const SPI_SETDOUBLECLKWIDTH As Integer = 29
    ' setzt die maximale Weite, die sich der Mauszeiger bewegen darf, damit Windows einen
    ' Doppelklick erkennt. "uiParam" erwartet den neuen Wert und "pvParam" muss `0` sein.

    Const SPI_SETDRAGFULLWINDOWS As Integer = 37
    ' (nur Win 9x) setzt Informationen, ob der Fensterinhalt beim Verschieben oder Verändern der Größe
    ' eines Fenster angezeigt (ungleich 0) oder versteckt (0) werden soll. "uiParam"
    ' erwartet den neuen Wert und "pvParam" muss `0` sein.

    Const SPI_SETDRAGHEIGHT As Integer = 77
    ' (nur Win 9x) setzt die maximale Höhe, die sich der Mauscursor bewegen muss, damit
    ' Windows eine Drag & Drop-Operation erkennt. "uiParam" erwartet den neuen Wert und
    ' "pvParam" muss `0` sein.

    Const SPI_SETDRAGWIDTH As Integer = 76
    ' (nur Win 9x) setzt die maximale Weite, die sich der Mauscursor bewegen muss, damit
    ' Windows eine Drag & Drop-Operation erkennt. "uiParam" erwartet den neuen Wert und
    ' "pvParam" muss `0` sein.

    Const SPI_SETFASTTASKSWITCH As Integer = 36
    ' schaltet das Wechseln der Fenster per ALT + TAB an (ungleich 0) oder aus (0).
    ' "uiParam" erwartet den neuen Wert und "pvParam" muss `0` sein.

    Const SPI_SETFILTERKEYS As Integer = 51
    ' setzt Eigenschaften der FILTERKEYS die z.B. für die Wiederholrate der Tasten der
    ' Tastatur verantwortlich sind. "uiParam" erwartet die Größe der FILTERKEYS-Struktur,
    ' die bei "pvParam" übergeben werden muss.

    Const SPI_SETFONTSMOOTHING As Integer = 75
    ' schaltet das Weichzeichnen der Systemfonts ein (ungleich 0) oder aus (0).
    ' "uiParam" erwartet den neuen Wert und "pvParam" muss `0` sein.

    Const SPI_SETGRIDGRANULARITY As Integer = 19
    ' setzt die Körnung des Desktopgitters. "uiParam" erwartet den neuen Wert und
    ' "pvParam" muss `0` sein.

    Const SPI_SETHIGHCONTRAST As Integer = 67
    ' (nur Win 9x) setzt die HIGHCONTRAST-Eigenschaften, die die Anzeige der Fenster
    ' steuern. "uiParam" muss `0` sein und "pvParam" erwartet die HIGHCONTRAST-Struktur.

    Const SPI_SETICONMETRICS As Integer = 46
    ' (nur Win 9x) setzt die Eigenschaften, wie Windows die Icons anzeigt. "uiParam" muss `0`
    ' sein und "pvParam" erwartet die ICONMETRICS-Struktur.

    Const SPI_SETICONS As Integer = 88
    ' lädt die System-Icons neu. "uiParam" und "pvParam" müssen `0` sein.

    Const SPI_SETICONTITLELOGFONT As Integer = 34
    ' setzt die Schriftart für die Icon-Titeltexte. "uiParam" erwartet die Größe der
    ' LOGFONT-Struktur, die bei "pvParam" übergeben werden muss.

    Const SPI_SETICONTITLEWRAP As Integer = 26
    ' setzt Informationen, ob Icon-Titeltexte in einer Zeile (0) angezeigt werden sollen oder ob sie bei
    ' entsprechender Länge auf mehrere Zeilen (ungleich 0) verteilt werden sollen.
    '  "uiParam" erwartet den Wert, um die Eigenschaft festzulegen und "pvParam" muss `0` sein.

    Const SPI_SETKEYBOARDDELAY As Integer = 23
    ' setzt die Zeit, die zwischen dem ersten Tastendruck einer Tastaturtaste und
    ' dem Einsetzen der Wiederholrate in Sekunden vergeht. "uiParam" erwartet einen Wert
    ' zwischen `0` und `3`. "pvParam" muss `0` sein.

    Const SPI_SETKEYBOARDPREF As Integer = 69
    ' (nur Win 9x) setzt Informationen, dass Windows das System tastaturfreundlich auslegen soll,
    ' weil es das Haupteingabegerät ist, wenn bei "uiParam" eine Zahl `ungleich 0` übergeben
    ' wird. Andernfalls wird diese Eigenschaft nicht gesetzt.

    Const SPI_SETKEYBOARDSPEED As Integer = 11
    ' setzt die Wiederholungsrate der Tastatur. "uiParam" erwartet einen Wert zwischen _
    '' `0` und `31`. "pvParam" muss `0` sein.

    Const SPI_SETLANGTOGGLE As Integer = 91
    ' (nur Win 9x) setzt Informationen, ob ein Hotkey benutzt werden soll, um die Tastatursprache zu
    ' ändern, der neue Hotkey wird aus der Registry gelesen unter
    ' "HKEY_CURRENT_USER\keyboard layout\toggle". "uiParam"und "pvParam" müssen `0` sein.

    Const SPI_SETLOWPOWERACTIVE As Integer = 85
    ' (nur Win 9x) setzt Informationen, ob der LowPower Modus aktiviert (1) oder deaktiviert (0) sein
    ' soll. "uiParam" erwartet den neuen Wert und "pvParam" muss `0` sein.

    Const SPI_SETLOWPOWERTIMEOUT As Integer = 81
    ' (nur Win 9x) setzt die Zeit ,die vergehen muss, bis der LowPower-Modus aktiviert
    ' wird. "uiParam" erwartet den neuen Wert in Sekunden und "pvParam muss `0` sein.

    Const SPI_SETMENUDROPALIGNMENT As Integer = 28
    ' setzt Informationen, ob die Popupmenüs links (0) oder rechts (ungleich 0) auftauchen sollen.
    ' "uiParam" erwartet den neuen Wert und "pvParam muss `0` sein.

    Const SPI_SETMINIMIZEDMETRICS As Integer = 44
    ' setzt die Eigenschaften, wie minimierte Fenster angezeigt werden. "uiParam" muss
    ' `0` sein und "pvParam" erwartet eine gefüllte MINIMIZEDMETRICS-Struktur.

    Const SPI_SETMOUSE As Integer = 4
    ' setzt die Eigenschaften der X- & Y-Achse und die Mausgeschwindigkeit. "uiParam"
    ' muss `0` sein und "pvParam" erwartet ein Long-Array mit 3 Feldern, die
    ' nacheinander die X-Achsen, Y-Achsen und Mausgeschwindigkeits-Eigenschaften
    ' enthalten muss.

    Const SPI_SETMOUSEBUTTONSWAP As Integer = 33
    ' setzt Informationen, dass die Mausbuttons getauscht werden. Übergeben Sie bei "uiParam"
    ' den Wert  `0` für die Originaleinstellung oder `ungleich 0` für vertauschte Mausbuttons.
    ' "pvParam" muss `0` sein.

    Const SPI_SETMOUSEKEYS As Integer = 55
    ' Setzt die MOUSKEYS-Eigenschaften, die es ermöglichen, die Maus per Keypad zu
    ' steuern. "uiParam" erwartet die Größe der MOUSEKEYS-Struktur, die bei "pvParam"
    ' gefüllt übergeben werden muss.

    Const SPI_SETMOUSETRAILS As Integer = 93
    ' (nur Win 9x) setzt die Länge der Mausspur. Wird bei "uiParam" der Wert `1` oder `0`
    ' übergeben, so ist die Mausspur aus, alle größeren Werte symbolisieren die Anzahl
    ' der anzuzeigenden Mauscursor der Mausspur. "pvParam" muss `0` sein.

    Const SPI_SETNONCLIENTMETRICS As Integer = 42
    ' (nur Win 9x) setzt die Eigenschaften des Clientbereiches eines MDI-Fensters.
    ' "uiParam" muss `0` sein und "pvParam" erwartet eine NONCLIENTMETRICS-Struktur.

    Const SPI_SETPENWINDOWS As Integer = 49
    ' (Nur Win 9x) lädt (ungleich 0) oder entlädt (0) "Pen für Windows". "uiParam"
    ' erwartet den Wert zum Entladen oder Laden des Programms und "pvParam" muss `0` sein.

    Const SPI_SETPOWEROFFACTIVE As Integer = 86
    ' (nur Win 9x) setzt Informationen, ob das Betriebssystem sich nach einer bestimmten Zeit in den
    ' PowerOff-Modus (ungleich 0) schaltet oder nicht (0). "uiParam" erwartet den Wert
    ' der den PowerOff-Modus aktiviert oder nicht und "pvParam" muss `0` sein.

    Const SPI_SETPOWEROFFTIMEOUT As Integer = 82
    ' (nur Win 9x) setzt die Zeit, die das System auf eine Benutzereingabe wartet, bevor
    ' der PowerOff-Modus gestartet wird. "uiParam" erwartet die Anzahl an Sekunden, die
    ' das System warten soll. "pvParam" muss `0` sein.

    Const SPI_SETSCREENREADER As Integer = 71
    ' (nur Win 9x) setzt Informationen, ob ein ScreenReader-Programm läuft (ungleich 0) oder nicht (0).
    ' "uiParam" erwartet den Wert der besagt, ob ein ScreenReader ausgeführt wird und
    ' "pvParam" muss `0` sein.

    Const SPI_SETSCREENSAVEACTIVE As Integer = 17
    ' setzt Informationen ob der Bildschirmschoner nach einer bestimmten Anzahl von Sekunden gestartet
    ' werden soll (unglecih 0) oder nicht (0). "uiParam" erwartet den Wert zum Aktivieren
    ' oder Deaktivieren des Bildschirmschoners und "pvParam" muss `0` sein.

    Const SPI_SETSCREENSAVETIMEOUT As Integer = 15
    ' setzt das Zeitlimit, nach dem der Bildschirmschoner gestartet werden soll. "uiParam"
    ' erwartet die Anzahl an Sekunden, nach dem der Bildschirmschoner gestartet wird und
    ' "pvParam" muss `0` sein.

    Const SPI_SETSERIALKEYS As Integer = 63
    ' (nur Win 9x) setzt die Eigenschaften der SERIALKEYS, die die Ports steuern.
    ' "uiParam" muss `0` sein und "pvParam" erwartet eine SERIALKEYS-Struktur.

    Const SPI_SETSHOWSOUNDS As Integer = 57
    ' setzt Informationen, ob SHOWSOUNDS aktiviert (ungleich 0) oder deaktiviert (0) sind. "uiParam"
    ' erwartet den Wert zum Deaktivieren oder Aktivieren der SHOWSOUNDS und "pvParam"
    ' muss `0` sein.

    Const SPI_SETSOUNDSENTRY As Integer = 65
    ' setzt die Eigenschaften der SOUNDENTRY-Eigenschaften, die akustische Signale mit
    ' visuellen Effekten untermalen. "uiParam" erwartet die Größe der SOUNDENTRY-Struktur,
    ' die bei "pvParam" gefüllt übergeben werden muss.

    Const SPI_SETSTICKYKEYS As Integer = 59
    ' setzt die STICKYKEYS, die die Erkennung von dem gleichzeitigen Drücken von zwei
    ' Tastaturtasten einrichten um so z.B. eine zeitliche Differenz auszubessern.
    ' "uiParam" erwartet die Größe der STICKYKEYS-Struktur, die bei "pcParam" gefüllt übergeben
    ' werden muss.

    Const SPI_SETTOGGLEKEYS As Integer = 53
    ' setzt die Eigenschaften der TOGGELKEYS, die Sounds abspielen wenn die Num-, die
    ' Feststell- oder Rollen-Taste gedrückt wird. "uiParam" erwartet die Größe der
    ' TOGGELKEYS-Struktur, die bei "pvParam" gefüllt übergeben werden muss.

    Const SPI_SETWORKAREA As Integer = 47
    ' (nur Win 9x) setzt den Arbeitsbereich des Desktops fest. "uiParam" muss `0` sein und
    ' "pvParam" erwartet eine gefüllte RECT-Struktur.

    Const SPI_GETMENUUNDERLINES As Integer = &H100A
    Const SPI_SETMENUUNDERLINES As Integer = &H100B

    Const SPIF_SENDWININICHANGE As Integer = &H2
    Const SPIF_UPDATEINIFILE As Integer = &H1

'Declare Function WNetAdd Lib "mpr.dll" Alias "WNetAddConnectionA" ( _
  ByVal NetworkPath$, ByVal Password$, ByVal LocalName$) As Long
'Declare Function WNetCancel Lib "mpr.dll" Alias "WNetCancelConnectionA" ( _
'  ByVal LocalName$, ByVal bForce&) As Long
Declare Function WNetGet& Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName$, ByVal lpszRemoteName$, cbRemoteName&)
Declare Function GetDriveType& Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive$)

' Benötigte API-Deklarationen
Public Const NO_ERROR = 0
Public Const CONNECT_UPDATE_PROFILE = &H1
' The following includes all the constants defined for NETRESOURCE,
' not just the ones used in this example.
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_REMEMBERED = &H3
Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
' Error Constants:
Public Const ERROR_Access_DENIED = 5&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_BAD_DEV_TYPE = 66&
Public Const ERROR_BAD_DEVICE = 1200&
Public Const ERROR_BAD_NET_NAME = 67&
Public Const ERROR_BAD_PROFILE = 1206&
Public Const ERROR_BAD_PROVIDER = 1204&
Public Const ERROR_BUSY = 170&
Public Const ERROR_CANCELLED = 1223&
Public Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Public Const ERROR_EXTENDED_ERROR = 1208&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&

Private Type NETRESOURCE
  dwScope As Long
  dwType As Long
  dwDisplayType As Long
  dwUsage As Long
  lpLocalName As String
  lpRemoteName As String
  lpComment As String
  lpProvider As String
End Type

Declare Function WNetAddConnection2& Lib "mpr.dll" Alias "WNetAddConnection2A" _
                 (lpNetResource As NETRESOURCE, ByVal lpPassword$, ByVal lpUserName$, ByVal dwFlags&)
Declare Function WNetCancelConnection2& Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName$, ByVal dwFlags&, ByVal fForce&)
Declare Function GetLogicalDriveStrings& Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength&, ByVal lpBuffer$)
Declare Function GetComputerName& Lib "kernel32" Alias "GetComputerNameA" (ByVal lbbuffer$, nSize&)

Declare Function FindWindowEx& Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent&, ByVal hwndChildAfter&, _
  ByVal lpszClass$, _
  ByVal lpszWindow$)
Declare Function GetDesktopWindow& Lib "user32" ()
Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Declare Function GetWindow& Lib "user32" (ByVal hwnd&, ByVal wCmd&)
'Declare Function ShowWindow& Lib "User32" (ByVal hwnd&, ByVal nCmdShow&)
'Declare Function WindowFromPoint& Lib "User32" (ByVal xPoint&, ByVal yPoint&)

Declare Function SendMessage& Lib "user32" Alias "SendMessageA" _
                (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
    
' Im Modul
Public Const NERR_SUCCESS As Long = 0&

' Freigabe Typen
Const STYPE_ALL       As Long = -1
Const STYPE_DISKTREE  As Long = 0
Const STYPE_PRINTQ    As Long = 1
Const STYPE_DEVICE    As Long = 2
Const STYPE_IPC       As Long = 3
Const STYPE_SPECIAL   As Long = &H80000000

' Rechte
Const ACCESS_READ     As Long = &H1
Const ACCESS_WRITE    As Long = &H2
Const ACCESS_CREATE   As Long = &H4
Const ACCESS_EXEC     As Long = &H8
Const ACCESS_DELETE   As Long = &H10
Const ACCESS_ATRIB    As Long = &H20
Const ACCESS_PERM     As Long = &H40
Const ACCESS_ALL      As Long = ACCESS_READ Or _
                                ACCESS_WRITE Or _
                                ACCESS_CREATE Or _
                                ACCESS_EXEC Or _
                                ACCESS_DELETE Or _
                                ACCESS_ATRIB Or _
                                ACCESS_PERM

Type SHARE_INFO_2
  shi2_netname       As Long
  shi2_type          As Long
  shi2_remark        As Long
  shi2_permissions   As Long
  shi2_max_uses      As Long
  shi2_current_uses  As Long
  shi2_path          As Long
  shi2_passwd        As Long
End Type
  

Declare Function NetShareAdd& Lib "netapi32" (ByVal servername$, ByVal level&, buf As Any, parmerr&)

Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = &HFF&
Private Const LANG_USER_DEFAULT             As Long = &H400&

Declare Function FormatMessage& Lib "kernel32" Alias "FormatMessageA" _
  (ByVal dwFlags&, ByRef lpSource As Any, ByVal dwMessageId&, ByVal dwLanguageId&, _
   ByVal lpBuffer$, ByVal nSize&, ByRef Arguments&)

Public EigDatAnmL$ ' = "D:\daten\eigene Dateien\" ' "E:\eigene Dateien alt" ' für Notbetrieb
Public DatenAnmL$ ' = "d:\daten" auf anmeldl
Public PatDokAnmL$ ' = "D:\daten\Patientendokumente\" ' "E:\P"   ' für Notbetrieb
Public DownAnmL$ ' = "D:\daten\down" ' "E:\down"
'Public Const DownServer$ = "D:\down"
Public GeraldAnmL$ ' = "d:\shome\gerald"
Public KothnyAnmL$ ' = "d:\shome\kothny"
Public ReadOAnmL$ ' = "d:\shome\gerald"
Public ReadOKAnmL$ ' = "d:\shome\kothny"
Dim LWTrekstor$
Dim alBoot$, alVolume$, alData$
Dim arRecover$, arData$, arBackup$ ' , TMServCptServer$ ' = "H:\turbomed"
Dim mitteRoot$, mitteVol$, mitteAustausch$ 'c:, i:, o:
Dim anmeldlBoot$, anmeldlVol$, anmeldlData$
Dim anmeldrZweit$, anmeldrBackup$
Dim sonoBoot$, sonoDaten$
Public alDasi$ ' "D:\Turbomed-Dasi"
Public arDasi$ ' "E:\Turbomed-Dasi"
Public arDokumente$  ' "H:\turbomed\Dokumente"
'Public Const DSiServer$ = "E:\Turbomed-DASI"

Public sysdrv$
Dim POk%
Public PatDok$ ' "\\linux1\Gemein\patdok" / \\linmitte\sam\p "\\mitte\p"
Dim EigDat$ ' "\\linux1\daten\eigene Dateien" / "\\linmitte\eigene Dateien" "\\MITTE\u"
Dim Gerald$ ' "\\linux1\daten\shome\gerald" \\linmitte\sam\gerald
Dim Kothny$ ' "\\linux1\daten\shome\kothny" \\linmitte\sam\kothny
Dim ReadO$ ' "\\linux1\geraldprivat" \\linmitte\sam\geraldprivat
Dim ReadOK$ ' "\\linux1\kothnyprivat" \\linmitte\sam\kothnyprivat
Dim Down$ ' "\\linux1\daten\down" / "\\MITTE\v" \\linmitte\down
Public TMStammV$ ' "\\linux1\turbomed" / \\linmitte\turbomed "\\ANMELDR\Turbomed"
Dim TMStammVk$ ' "\turbomed" / \\tm"
Dim Programme$ ' "\\linux1\daten\Programme",\\linmitte\sam\programme "\\MITTE\Gemein\Programme"
Dim Sicherheit$ ' "\\ANMELDR\Sicherheit"
Dim Dokumente$ ' "\\linux1\Gemein\Dokumente" \\linmitte\DAT\turbomed\dokumente
Dim TMServCpt$ ' "linux1", "linmitte", "ANMELDR"
Dim DSi$ ' \\ANMELDL\TM-DASI, D:\Turbomed-DASI ' \\ANMELDR\TM-DASI, e:\Turbomed-DASI
Public obNot% ' ob Notbetrieb
' Folgende entsprechen immer außer im Notbetrieb auf MITTE bzw. ANMELDR den Systemeinstellungen
Public EigDatDirekt$, PatDokDirekt$, DownDirekt$, ReadODirekt$, ReadOKDirekt$, GeraldDirekt$, KothnyDirekt$, TMServCptDirekt$, DokumenteDirekt$
Public Cpt$, UN$, AUP$, userprof$, TMExeV$, TMNotV$, TMNotPr$, lokalTMExeV$, idt As TMIniDatei
Dim autoVz$, StartMen$, StartMenProg$, Favor$
Public FSO As New FileSystemObject ' %windir%\system32\scrrun.dll ' as object
'Public wsh As IWshShell_Class ' %windir%\system32\wshom.ocx
Public wsh2 As IWshRuntimeLibrary.WshShell

' Folgendes zum Ermitteln aller Benutzer
' benötigte API-Deklarationen
Private Declare Function NetApiBufferFree& Lib "netapi32.dll" (ByVal lpBuffer&)
Private Declare Function NetUserEnum& Lib "netapi32.dll" (servername As Byte, ByVal level&, ByVal filter&, bufptr&, ByVal prefmaxlen&, entriesread&, totalentries&, resume_handle&)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length&)
Private Declare Function lstrlen& Lib "kernel32" Alias "lstrlenA" (ByVal lpString$)
'Public WMIreg As SWbemObjectEx  ' %windir%\system32\wbem\wbemdisp.tlb
'Public WMIreg As WbemScripting.SWbemObjectEx
'Public colItems As SWbemObjectSet
'Public objItem As SWbemObject
Public Declare Function GetShortPathName& Lib "kernel32" Alias "GetShortPathNameA" _
      (ByVal lpszLongPath$, ByVal lpszShortPath$, ByVal cchBuffer&)

' für Ping

'Icmp constants converted from
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_pingstatus.asp

Private Const ICMP_SUCCESS As Long = 0
Private Const ICMP_STATUS_BUFFER_TO_SMALL = 11001                   'Buffer Too Small
Private Const ICMP_STATUS_DESTINATION_NET_UNREACH = 11002           'Destination Net Unreachable
Private Const ICMP_STATUS_DESTINATION_HOST_UNREACH = 11003          'Destination Host Unreachable
Private Const ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH = 11004      'Destination Protocol Unreachable
Private Const ICMP_STATUS_DESTINATION_PORT_UNREACH = 11005          'Destination Port Unreachable
Private Const ICMP_STATUS_NO_RESOURCE = 11006                       'No Resources
Private Const ICMP_STATUS_BAD_OPTION = 11007                        'Bad Option
Private Const ICMP_STATUS_HARDWARE_ERROR = 11008                    'Hardware Error
Private Const ICMP_STATUS_LARGE_PACKET = 11009                      'Packet Too Big
Private Const ICMP_STATUS_REQUEST_TIMED_OUT = 11010                 'Request Timed Out
Private Const ICMP_STATUS_BAD_REQUEST = 11011                       'Bad Request
Private Const ICMP_STATUS_BAD_ROUTE = 11012                         'Bad Route
Private Const ICMP_STATUS_TTL_EXPIRED_TRANSIT = 11013               'TimeToLive Expired Transit
Private Const ICMP_STATUS_TTL_EXPIRED_REASSEMBLY = 11014            'TimeToLive Expired Reassembly
Private Const ICMP_STATUS_PARAMETER = 11015                         'Parameter Problem
Private Const ICMP_STATUS_SOURCE_QUENCH = 11016                     'Source Quench
Private Const ICMP_STATUS_OPTION_TOO_BIG = 11017                    'Option Too Big
Private Const ICMP_STATUS_BAD_DESTINATION = 11018                   'Bad Destination
Private Const ICMP_STATUS_NEGOTIATING_IPSEC = 11032                 'Negotiating IPSEC
Private Const ICMP_STATUS_GENERAL_FAILURE = 11050                   'General Failure

Public Const WINSOCK_ERROR = "Windows Sockets not responding correctly."
Public Const INADDR_NONE As Long = &HFFFFFFFF
Public Const WSA_SUCCESS = 0
Public Const WS_VERSION_REQD As Long = &H101

'Clean up sockets.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512

Private Declare Function WSACleanup Lib "wsock32.dll" () As Long

'Open the socket connection.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long

'Create a handle on which Internet Control Message Protocol (ICMP) requests can be issued.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpcreatefile.asp
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

'Convert a string that contains an (Ipv4) Internet Protocol dotted address into a correct address.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winsock/wsapiref_4esy.asp
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

'Close an Internet Control Message Protocol (ICMP) handle that IcmpCreateFile opens.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpclosehandle.asp

Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long

'Information about the Windows Sockets implementation
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To 256) As Byte
   szSystemStatus(0 To 128) As Byte
   iMaxSockets As Long
   iMaxUdpDg As Long
   lpVendorInfo As Long
End Type

'Send an Internet Control Message Protocol (ICMP) echo request, and then return one or more replies.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIcmpSendEcho.asp
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
 
'This structure describes the options that will be included in the header of an IP packet.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIP_OPTION_INFORMATION.asp
Private Type IP_OPTION_INFORMATION
   Ttl             As Byte
   Tos             As Byte
   flags           As Byte
   OptionsSize     As Byte
   OptionsData     As Long
End Type

'This structure describes the data that is returned in response to an echo request.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmp_echo_reply.asp
Public Type ICMP_ECHO_REPLY
   address         As Long
   Status          As Long
   RoundTripTime   As Long
   DataSize        As Long
   Reserved        As Integer
   ptrData                 As Long
   Options        As IP_OPTION_INFORMATION
   data            As String * 250
End Type ' ICMP_ECHO_REPLY





Function sayifisIDE()
 IsIDE = True
End Function ' sayifisIDE

Sub meßzeit()
 Static T1#, T2#
 If T1 = T2 Then
  T1 = Timer
 Else
  T2 = Timer
 End If
 ' ' fi.Laufzeit = "Laufzeit:    " & CStr(T2 - T1) & "s"
End Sub ' meßzeit

Sub BenutzerWMI()
' Beginn
'strSearch = InputBox("Zu welchem Namen wird der SID gesucht?")
 Dim objWMI, strWQL$, objResult, objAcc, strResult
 Set objWMI = GetObject("winmgmts:")
 strWQL = "select * from win32_account where sidtype=1" ' where Name='" & strSearch & "'"
 Set objResult = objWMI.ExecQuery(strWQL)
 
 For Each objAcc In objResult
  Debug.Print objAcc.SID, objAcc.name, objAcc.sidtype
 Next
End Sub ' BenutzerWMI

Public Function EnumUsers(Optional ByVal sComputer As String = "") As Variant
  ' alle im System eingerichteten Usern ermitteln
  ' Rückgabe als String-Array
  Dim bServer() As Byte
  Dim nUsers() As Long
  Dim nBufPtr As Long
  Dim nCount As Long
  Dim nTotal As Long
  Dim i As Long, pos&
  Dim sUsers() As New CString
  Dim nBuffer() As Byte
  On Error GoTo fehler
  ' Computername (Server)
  ' (wird als Byte-Array benötigt)
  If LenB(sComputer) = 0 Then sComputer = Environ$("COMPUTERNAME")
  If Left$(sComputer, 2) <> "\\" Then sComputer = "\\" & sComputer
  bServer = sComputer & vbNullChar
  ' Benutzer ermitteln
  If NetUserEnum(bServer(0), 0, &H2, nBufPtr, 255&, nCount, nTotal, 0&) = 0 Then
    ReDim nUsers(nCount - 1)
    ReDim sUsers(nCount - 1)
    CopyMemory nUsers(0), ByVal nBufPtr, nCount * 4
    ' jetzt die Benutzernamen anhand der Speicheradresse
    ' (Pointer) ermitteln
    For i = 0 To nCount - 1
      If lstrlen(nUsers(i)) > 0 Then
        ReDim nBuffer(255)
        CopyMemory nBuffer(0), ByVal nUsers(i), 255
        sUsers(i) = nBuffer
'        If InStr(sUsers(i), vbNullChar) > 0 Then
'          sUsers(i) = Left$(sUsers(i), InStr(sUsers(i), vbNullChar) - 1)
        pos = sUsers(i).Instr(vbNullChar)
        If pos <> 0 Then
           sUsers(i).Cut (pos - 1)
        End If
      End If
    Next i
    ' Resourcen freigeben
    NetApiBufferFree nBufPtr
  End If
  EnumUsers = sUsers
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in EnumUsers/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' EnumUsers

#If Test Then
Function testUsers()
 Dim gname As New CString, ulist() As New CString, i&
 On Error GoTo fehler
 gname = "fjdkals schade"
 ulist = EnumUsers
 For i = 0 To UBound(ulist)
  If gname.Right(ulist(i).length) = ulist(i) Then Stop
 Next i
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in testUsers/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' testUsers
#End If

Function StrOhneUser(Str$) As CString
 Dim i&, ulist() As New CString
 On Error GoTo fehler
 Set StrOhneUser = New CString
 StrOhneUser = Str
 ulist = EnumUsers
 For i = 0 To UBound(ulist)
  If StrOhneUser.Right(ulist(i).length + 1) = " " & ulist(i) Then StrOhneUser.Cut (StrOhneUser.length - ulist(i).length - 1): Exit For
 Next i
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in strOhneUser/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' StrOhneUser

#If SchedulingAgentVB6 Then
Private Sub ScheduleJob(name$, Exe$, arg$, Comment$, WD$, uName$, pwd$)
    Dim jb As CJob
    Dim trg As CTrigger

    ' Create the Job
    Set jb = TaskScheduler.Jobs.Add(name)

    ' Set the job to run
    With jb

        ' The app to run
        .Application = Exe

        ' The command line (if your app needs command line arguments)
        .CommandLine = arg

        ' A comment
        .Comments = Comment

        ' Set the jobs working directory
        .WorkingDirectory = WD

        ' Required on NT/2000
        .SetAccountInfo uName, pwd

        ' If you don't want the user to see this job in the task scheduler
interface
        .Hidden = True

        ' Now add the time schedule
         Set trg = .Triggers.Add()
    End With

    With trg
        .MinutesInterval = 20
        .StartTime = #8:00:00 AM#
    End With
End Sub ' TaskReady = 267008
#End If

Private Sub cmdNew_Click(tName$, programm$)
Dim sTaskName As String
Dim oTask As Task
Dim m_oSchedule As New TaskScheduler2.Schedule
   On Error Resume Next
   If Err.Number = 0 Then
      On Error GoTo ErrHandler
GetTaskName:
      sTaskName = tName
      If StrPtr(sTaskName) Then
         Set oTask = m_oSchedule.CreateTask(sTaskName)
         On Error Resume Next
         With oTask
            ' Set the exe path
            .ApplicationName = programm
            ' Set other properties
            .Creator = "Task Scheduler Class sample"
            ' Set default FlagsText
            .flags = Interactive Or RunOnlyIfLoggedOn
            ' Add a m_oSchedule
            With .Triggers
               With .Add
                  .BeginDay = Date
                  .StartTime = Time
                  .TriggerType = DailyTrigger
                  .Daily_Interval = 1
               End With ' .Add
            End With ' .Triggers
            ' Set the account in which the
            ' application will run
            .SetAccountInfo Environ$("USERNAME"), vbNullString
            ' Save the task
            .Save
         End With ' oTask
         ' Refresh the task list
      End If ' StrPtr(sTaskName) Then
   End If ' Err.Number = 0 Then
   Exit Sub
ErrHandler:
   MsgBox "The task already exists. Please enter a new name", vbExclamation
   Resume GetTaskName
End Sub ' cmdNew_Click

Function XmlDuration( _
    Optional ByVal Years As Integer, _
    Optional ByVal Months As Integer, _
    Optional ByVal Days As Integer, _
    Optional ByVal Hours As Integer, _
    Optional ByVal Minutes As Integer, _
    Optional ByVal Seconds As Integer) As String
    'In theory values like "P0YT20M" are valid, but Task Scheduler
    'seems to reject them.  Use this strategy to suppres zeros.
    Dim strDate As String
    Dim strTime As String
    
    strDate = "P"
    If Years > 0 Then strDate = strDate & CStr(Years) & "Y"
    If Months > 0 Then strDate = strDate & CStr(Months) & "M"
    If Days > 0 Then strDate = strDate & CStr(Days) & "D"
    
    strTime = "T"
    If Hours > 0 Then strTime = strTime & CStr(Hours) & "H"
    If Minutes > 0 Then strTime = strTime & CStr(Minutes) & "M"
    If Seconds > 0 Then strTime = strTime & CStr(Seconds) & "S"
    
    If Len(strTime) = 1 Then strTime = ""
    XmlDuration = strDate & strTime
End Function ' XmlDuration

Function XmlTime(ByVal Timestamp As Date) As String
    XmlTime = Format$(Timestamp, "yyyy-mm-dd\THh:Nn:Ss")
End Function ' XmlTime

Sub machtask2(Desc$, Applic$, args$, Optional usr$, Optional pwd$)
 Static Fehlerzahl&, strID$, ErrDescr$
 Dim TaskDef As Object 'TaskScheduler.ITaskDefinition
 On Error GoTo fehler
 If WV >= win_vista Then
  Dim Datei$, Pfad$
  Pfad = Left$(Applic, InStrRev(Applic, "\"))
  Datei = Mid$(Applic, InStrRev(Applic, "\") + 1)
  Dim ntask
  Set ntask = CreateObject("Schedule.Service") 'New TaskScheduler.TaskScheduler
  With ntask
    .Connect
   'This method call forces us to use late binding here, because .NewTask()
   'takes an Unsupported Variant type argument "flags" (UInt?).  It has to
   'be 0 anyway, so this works fine:
   Set TaskDef = .NewTask(0)
  strID = Format(Now(), "yyyymmdd-hhMMSS") ' Replace$(txtLogFile.Text, ".", "-") 'Create an Id value, no periods here!
        With TaskDef
            With .RegistrationInfo
                .Description = Desc
                .Author = "Schade"
            End With
            With .Principal
                .id = "P" & strID
                'If chkS4U.Value = vbUnchecked Then
                .LogonType = 1 ' TASK_LOGON_PASSWORD
                'If chkRunLevelHighest.Value = vbChecked Then
                .RunLevel = 1 ' TASK_RUNLEVEL_HIGHEST
            End With
             With .Settings
                .Enabled = True
                .StartWhenAvailable = True
                .ExecutionTimeLimit = XmlDuration(Minutes:=20) 'Not sure why we have two of these vv
                .WakeToRun = 0 ' chkWake.Value = vbChecked
                .Priority = 7 ' THREAD_PRIORITY_BELOW_NORMAL 'Actually the default.
                .DisallowStartIfOnBatteries = True
                .StopIfGoingOnBatteries = False ' True
                
                'Task will be writing to a database and we don't want to risk corruption,
                'but CreateDB is Formless and won't be able to process events:
                .AllowHardTerminate = True
                
                .Hidden = False
            End With
            With .Triggers.Create(1) ' TASK_TRIGGER_TIME)
                .StartBoundary = XmlTime(Now()) ' dtpStart.Value)
                .EndBoundary = XmlTime(DateAdd("yyyy", 25, Now())) ' Dateadd("n",30,dtpStart.Value)) '30 minutes from dtpStart.
                .ExecutionTimeLimit = XmlDuration(Minutes:=20) 'Not sure why we have two of these ^^
                .id = "T" & strID
'                .repetition.Duration = "P99YM"
                .repetition.Interval = "PT2M"
                .Enabled = True
            End With
            With .Actions.Create(0) 'TASK_ACTION_EXEC)
                .Path = Applic
                .WorkingDirectory = Pfad
                .Arguments = args
            End With
        End With
        With .GetFolder("\") ' .GetFolder("\")
            Dim ohnepw%
            If LenB(usr) = 0 And LenB(pwd) = 0 Then ohnepw = True
            If ohnepw Then
                On Error Resume Next
                .RegisterTaskDefinition _
                    Desc, _
                    TaskDef, _
                    6, _
                    , _
                    , _
                    2
                    'TASK_CREATE_OR_UPDATE, _
                    'TASK_LOGON_S4U
                If Err Then
                    ErrDescr = Err.Description
                    rufauf App.Path & "\nachricht.exe", "Fehler beim Einplanen von: " & Applic & vbCrLf & ErrDescr, 0, , 0
                Else
'                    MsgBox "Task submitted", vbOKOnly, App.EXEName
                 End If
            Else
                On Error Resume Next
                .RegisterTaskDefinition _
                    Desc, _
                    TaskDef, _
                    6, _
                    usr, _
                    pwd, _
                    1 ' TASK_LOGON_PASSWORD
                If Err Then
                    ErrDescr = Err.Description
                    rufauf App.Path & "\nachricht.exe", "Fehler beim Einplanen von: " & Applic & vbCrLf & ErrDescr, 0, , 0
                Else
'                    MsgBox "Task submitted", vbOKOnly, App.EXEName
                End If
             End If
        End With
    End With
End If
Exit Sub
fehler:
 If Fehlerzahl = 0 Then
  ErrDescr = Err.Description
'  Call Shell(App.Path & "\nachricht.exe " & "Fehler beim Einplanen von: " & oT.ApplicationName & vbCrLf & ErrDescr)
'  SuSh App.Path & "\nachricht.exe " & "Fehler beim Einplanen von: " & oT.ApplicationName & vbCrLf & ErrDescr, 0, , 0, 1
  rufauf App.Path & "\nachricht.exe", "Fehler beim Einplanen von: " & Applic & vbCrLf & ErrDescr, 0, , 0
  Fehlerzahl = Fehlerzahl + 1
 End If
 Resume Next
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in machTask2/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' machtTask2

Sub machAufgb(tName$, AName$, CdL$, ByVal flags&, MRT&, Prio&, WD$, Comment$, ByVal Creator$, BegD As Date, i1&, i2&, I3&, Itv&, Dur&, StartT As Date, EndD As Date, TFlags As TriggerFlags, TriggerType As TriggerTypes, Optional Stat As JobStatus = 267008, Optional alsAdm%, Optional args$, Optional obAdm%, Optional obletztes%) ' TaskReady = 267008
 On Error GoTo fehler
 If WD = "" Then WD = ProgVerz
 If WV < win_vista Then
  xpmachTask tName, AName, CdL, flags, MRT, Prio, WD, Comment, Creator, BegD, i1, i2, I3, Itv, Dur, StartT, EndD, TFlags, TriggerType, Stat       ' TaskReady = 267008
'  domTgeht
 Else
 ' "Turbomed Ausfallwarnung", ProgVerz & "poetaktiv.exe", "", "administrator", AdminPwd
  w10machTask tName, AName, CdL, flags, MRT, Prio, WD, Comment, IIf(alsAdm, "Administrator", Creator), BegD, i1, i2, I3, Itv, Dur, StartT, EndD, TFlags, TriggerType, Stat, args, obAdm, obletztes ' TaskReady = 267008
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in machAufgb/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' machAufgb

'Sub domTgeht()
' Dim Sc As New TaskScheduler2.Schedule, erg&
' If LenB(Cpt) = 0 Then Cpt = CptName
' On Error Resume Next
' Sc.TargetComputer = "\\" & Cpt
' Sc.Refresh
' If InStrB(LCase$(Command), "task") <> 0 Then
'  ListTasks Sc
' End If
' On Error Resume Next
' xpmachTask "Turbomed Ausfallsystem", Environ("programfiles") & "\poetaktiv.exe", "", 0, 259200000, 32, Environ("programfiles"), vNS, "sturm", #6/24/2010#, 1, 0, 0, 2, 1439, #12:00:00 AM#, #11/30/1999#, 0, 1, 267011 ' "debug", am 10.11.11 entfernt
'End Sub

Sub w10machTask(tName$, AName$, CdL$, ByVal flags&, MRT&, Prio&, WD$, Comment$, ByVal Creator$, BegD As Date, i1&, i2&, I3&, Itv&, Dur&, StartT As Date, EndD As Date, TFlags&, TriggerType&, Optional Stat& = 267008, Optional args$, Optional obAdm%, Optional obletztes%) ' TaskReady = 267008
 Static cmd$
 Dim Verz$, strID$
 strID = Format(Now(), "yyyymmdd-hhMMSS")
 Verz = AppVerz & "\tasks"
 VerzPrüf Verz
 Open Verz & "\" & tName For Output As #113
 Print #113, "<?xml version=""1.0"" encoding=""UTF-16""?>"
 Print #113, "<Task version=""1.4"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">"
 Print #113, "  <RegistrationInfo>"
 Print #113, "    <Author>Schade</Author>"
 Print #113, "    <Description>" & tName & "</Description>"
 Print #113, "    <URI>\" & tName & "</URI>"
 Print #113, "  </RegistrationInfo>"
 Print #113, "  <Triggers>"
 Print #113, "    <TimeTrigger id=""T" & strID & """>"
 Print #113, "      <Repetition>"
 Print #113, "        <Interval>PT" & Itv & "M</Interval>"
 Print #113, "        <StopAtDurationEnd>" & IIf(TFlags And 2, "true", "false") & "</StopAtDurationEnd>"
 Print #113, "      </Repetition>"
 Print #113, "      <StartBoundary>" & Format(Now(), "yyyy-mm-ddThh:MM:SS") & "</StartBoundary>"
 Print #113, "      <EndBoundary>2060-09-02T16:19:28</EndBoundary>"
 Print #113, "      <ExecutionTimeLimit>PT" & (MRT / 60000) & "M</ExecutionTimeLimit>"
 Print #113, "      <Enabled>" & IIf(Not TFlags And 4, "true", "false") & "</Enabled>"
 Print #113, "    </TimeTrigger>"
 Print #113, "  </Triggers>"
 Print #113, "  <Principals>"
 Print #113, "    <Principal id=""P" & strID & """>"
 Print #113, "      <RunLevel>HighestAvailable</RunLevel>"
 Print #113, "      <UserId>" & Creator & "</UserId>"
 Print #113, "      <LogonType>Password</LogonType>"
 Print #113, "    </Principal>"
 Print #113, "  </Principals>"
 Print #113, "  <Settings>"
 Print #113, "    <MultipleInstancesPolicy>" & IIf(TFlags And 2, "StopExisting", "Queue") & "</MultipleInstancesPolicy>"
 Print #113, "    <DisallowStartIfOnBatteries>" & IIf(flags And 64, "true", "false") & "</DisallowStartIfOnBatteries>"
 Print #113, "    <StopIfGoingOnBatteries>" & IIf(flags And 128, "true", "false") & "</StopIfGoingOnBatteries>"
 Print #113, "    <AllowHardTerminate>" & IIf(flags And 1, "true", "false") & "</AllowHardTerminate>"
 Print #113, "    <StartWhenAvailable>" & IIf(flags And 8192, "false", "true") & "</StartWhenAvailable>"
 Print #113, "    <RunOnlyIfNetworkAvailable>" & IIf(flags And 1024, "true", "false") & "</RunOnlyIfNetworkAvailable>"
 Print #113, "    <IdleSettings>"
 Print #113, "      <StopOnIdleEnd>" & IIf(flags And 32, "false", "true") & "</StopOnIdleEnd>"
 Print #113, "      <RestartOnIdle>" & IIf(flags And 2048, "true", "false") & "</RestartOnIdle>"
 Print #113, "    </IdleSettings>"
 Print #113, "    <AllowStartOnDemand>true</AllowStartOnDemand>"
 Print #113, "    <Enabled>" & IIf(flags And 4, "false", "true") & "</Enabled>"
 Print #113, "    <Hidden>" & IIf(flags And 512, "true", "false") & "</Hidden>"
 Print #113, "    <RunOnlyIfIdle>" & IIf(flags And 16, "true", "false") & "</RunOnlyIfIdle>"
 Print #113, "    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>"
 Print #113, "    <UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine>"
 Print #113, "    <WakeToRun>" & IIf(flags And 4096, "true", "false") & "</WakeToRun>"
 Print #113, "    <ExecutionTimeLimit>PT" & (MRT / 60000) & "M</ExecutionTimeLimit>"
 Print #113, "    <Priority>7</Priority>"
 Print #113, "  </Settings>"
 Print #113, "  <Actions Context=""P" & strID & """>"
 Print #113, "    <Exec>"
 Print #113, "      <Command>" & AName & IIf(CdL = "", "", " ") & CdL & "</Command>"
 Print #113, "      <Arguments>" & args & "</Arguments>"
 Print #113, "      <WorkingDirectory>" & WD & "</WorkingDirectory>"
 Print #113, "    </Exec>"
 Print #113, "  </Actions>"
 Print #113, "</Task>"
 Close #113
 'rufauf "cmd", "/c move """ & Verz & "\" & tName & """ """ & Environ("windir") & "\system32\tasks\" & """", 2, , -1, 0
 Dim mitloe%
 mitloe = True
 If mitloe Then
  cmd = IIf(cmd = "", "", cmd & " & ")
  cmd = cmd & "schtasks /query /tn """ & tName & """ >nul && schtasks /delete /tn """ & tName & """ /f "
 End If
 cmd = IIf(cmd = "", "", cmd & " & ")
 cmd = cmd & "schtasks /query /tn """ & tName & """ >NUL 2>&1 || schtasks /create /xml """ & Verz & "\" & tName & """ /tn """ & tName & """ /ru " & IIf(obAdm, "administrator /rp " & AdminPwd, "system")
 If obletztes Then
  Dim obforce%
  obforce = True
  rufauf "cmd", "/c " & cmd, IIf(obforce, 1, 2), , 0, 0
 End If
'  If obforce Then
'   Dim rerg$
'   rerg = fragReg(2, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\" & tName, "Id")
'   If rerg <> "" Then
'  If FSO.FileExists("c:\windows\system32\tasks\" & tName) Then
'    rufauf "cmd", "/c schtasks /query /tn """ & tName & """ >NUL 2>&1 || schtasks /create /xml """ & Verz & "\" & tName & """ /tn """ & tName & """ /ru " & IIf(obAdm, "administrator /rp " & AdminPwd, "system"), 1, , -1, 0
'   End If ' rerg<>""
'  Else
'   rufauf "cmd", "/c schtasks /query /tn """ & tName & """ >nul && schtasks /delete /tn """ & tName & """ /f & schtasks /query /tn """ & tName & """ >NUL 2>&1 || schtasks /create /xml """ & Verz & "\" & tName & """ /tn """ & tName & """ /ru " & IIf(obAdm, "administrator /rp " & AdminPwd, "system"), 2, , 0, 0
'   rufauf "cmd", "/c schtasks /query /tn """ & tName & """ >nul && schtasks /delete /tn """ & tName & """ /f", 2, , 0, 0
'   rufauf "cmd", "/c schtasks /query /tn """ & tName & """ >NUL 2>&1 || schtasks /create /xml """ & Verz & "\" & tName & """ /tn """ & tName & """ /ru " & IIf(obAdm, "administrator /rp " & AdminPwd, "system"), 2, , -1, 0
'  End If ' obforce
' End If ' mitLoe
' rufauf "cmd", "/c schtasks /query /tn """ & tName & """ >%appdata%\erg0.txt 2>&1", 0, , -1, 0
' rerg = fragReg(2, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\" & tName, "Id")
' If rerg = "" Then
' rufauf "cmd", "/c schtasks /query /tn """ & tName & """ >%appdata%\erg1.txt 2>&1", 0, , -1, 0
'  rerg = fragReg(2, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\" & tName, "Id")
'  If rerg = "" Then
' If Not FSO.FileExists("c:\windows\system32\tasks\" & tName) Then
'  End If
' End If
' rufauf "cmd", "/c schtasks /query /tn """ & tName & """ >%appdata%\erg2.txt 2>&1", 0, , -1, 0
' rufauf "cmd", "/c schtasks /create /xml """ & Verz & "\" & tName & """ /tn """ & tName & """ /ru " & IIf(obAdm, "administrator /rp " & AdminPwd, "system"), 2, , -1, 0
 Exit Sub ' w10machTask
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in w10machTask/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' w10machTask

Sub xpmachTask(tName$, AName$, CdL$, ByVal flags&, MRT&, Prio&, WD$, Comment$, ByVal Creator$, BegD As Date, i1&, i2&, I3&, Itv&, Dur&, StartT As Date, EndD As Date, TFlags As TriggerFlags, TriggerType As TriggerTypes, Optional Stat As JobStatus = 267008)  ' TaskReady = 267008
' mehrfache Zeitpläne verwenden: nicht angezeigt
' Flags:
' interactive=1, deletewhendone=2, disabled=4, startonlyifidle= 16, killonidleend=32,
' dontstartifonbatteries= 64, killifgoingonbatteries=128, hidden=512,
' runifconnectedtointernet=1024,restartonidleresume=2048,systemrequired=4096,runonlyifloggedon=8192
'         2 = Task löschen, wenn nicht erneut geplant
'        16 = erst nach folgender Leerlaufzeit starten (hier nicht abgefragt), falls der PC nicht so lang im Leerlauf ist, erneut versuchen für maximal (hier nicht abgefragt)
'        32 = beenden, wenn der PC nicht mehr im Leerlauf ist
'        64 = nicht bei Akkubetrieb
'       128 = beenden, sobald Akkubetrieb einsetzt
'      4096 = Computer zur Ausführung der Task reaktivieren
' MRT: maximale Laufzeit in ms, -1 = Beenden nach nicht angekreuzt
' i1, i2: 0 0 = einmalig/bei Systemstart, 1 0 = täglich 1 1 = monatlich Januar, 1 2176 = mon Aug+Dez,
'         1 4095 = alle Mon
' i1: jedes i1.te Intervall (z.B. Tag, Woche) ausführen
' i2: 0 = einmalig, 1 = täglich, 2 = wöchentlich, 3 = Januar+Februar, 2176 = August+Dezember, 4095 = monatlich,
' itv: Wiederholungsintervall in Minuten
' TriggerType: 0 = normal 5 = im Leerlauf nach x Minuten (x nicht dabei), 6 = Systemstart, 7 = bei der Anmeldung
' Triggertype: Atlogon: 7, atsystemstart 6, dailytrigger 1, monthlydate 3, monthlyweekday 4, once 0,
'              OnIdle 5,WEEKLYTRIGGER 2
' i1 i2 i3 = 32 2176 1: jeden ersten Freitag im August und Dezember, tt = 4
' i1 i2 i3 = 1024 2176 0: jeden 11.8. und 11.12., tt = 3
'            9 26 0: jede 9. Woche am Mo Mi Do, tt = 2
'             7 0 0 = täglich jeden 7. Tag, tt = 1
'             1 x 0 = wöchentlich x=Mo 2, Di 4, Mi 8, Do 16, Fr 32, Sa 64, So 128, tt=2
'             0 0 0 = bei der Anmeldung, tt = 7
'             0 0 0 = im Leerlauf nach 13 min tt = 5
'             0 0 0 = bei Systemstart tt = 6
'  TFlags : hasenddate 1, killatdurationend 2, triggerdisabled = 4
'            2 = Task beenden, falls noch ausgeführt
 Dim oTask As Task, uName$, gname As New CString
 Static Fehlerzahl&, pwd$
 Static Sc As New TaskScheduler2.Schedule, scinit%
 On Error GoTo fehler
 If Not scinit Then
  On Error Resume Next
  Sc.TargetComputer = "\\" & Cpt
  If InStrB(LCase$(Command), "task") <> 0 Then
   ListTasks Sc
   ProgEnde
  End If
  On Error GoTo fehler
  scinit = True
 End If

 If StrPtr(tName) Then
   uName = Environ("username")
   Sc.Refresh
   gname = tName
   Select Case Creator
    Case "MySQL Administrator", "SYSTEM" ' Backups AppleUpdate,GoogleSoftwareUpdate
    Case Else
     If gname.Instr("Update") = 0 And gname <> "SystemIdleDetector" Then ' z.B. Google-Updater
      gname.AppVar Array(" ", uName)
     End If
   End Select
   Dim sci& ' 24.11.10 Task vorher löschen
   sci = 1
   Do
    On Error Resume Next
    Set oTask = Sc(sci)
    If Err.Number <> 0 Then Exit Do
    On Error GoTo fehler
    Debug.Print oTask.name
    If LCase$(oTask.name) = LCase$(gname) Then
     Sc.Delete (sci)
    End If
    sci = sci + 1
   Loop
 
   Set oTask = Sc.CreateTask(gname) ' cm_oSchedule.CreateTask(gname)
   On Error GoTo fehler
   With oTask
  ' Set the exe path
    .ApplicationName = AName
    .CommandLine = CdL
    If Stat = TaskDisabled And ((flags And 4) = 0) Then flags = flags + 4 Else If Stat = TaskReady And ((flags And 4) = 4) Then flags = flags - 4
    .flags = flags
    .MaxRunTime = MRT
    .Priority = Prio
    .WorkingDirectory = WD
    If Comment <> vNS Then .Comment = Comment
    ' Set other properties
    If LenB(Creator) = 0 Then .Creator = Environ("username") Else .Creator = Creator
    .Creator = Creator
    ' Set default FlagsText
'    .Flags = Interactive Or RunOnlyIfLoggedOn
    ' Add a m_oSchedule
    With oTask.Triggers.Add
     .BeginDay = BegD
     .TriggerType = TriggerType
      Select Case .TriggerType
       Case DailyTrigger
         .Daily_Interval = i1
       Case WeeklyTrigger: .Weekly_Interval = i1: .Weekly_DaysOfTheWeek = i2
       Case MONTHLYDATE: .MonthlyDate_Day = i1: .MonthlyDate_Months = i2
       Case MonthlyWeekDay: .MonthlyDOW_DaysOfTheWeek = i1: .MonthlyDOW_Months = i2: .MonthlyDOW_Week = I3
      End Select
     .Interval = Itv
     .Duration = Dur
     .StartTime = StartT
'   .Text = Text ' nur lesen
     .EndDay = EndD
     .flags = TFlags
    End With
    ' Set the account in which the
    ' application will run
    If InStrB(uName, "schade") <> 0 Or InStrB(uName, "gerald") <> 0 Then pwd = holap(uName) Else pwd = AdminPwd
'    .SetAccountInfo uname, vbNullString
    .SetAccountInfo LCase$(uName), pwd
    ' Save the task
    .Save
   End With
   ' Refresh the task list
   Sc.Refresh
  End If
  
 Exit Sub
fehler:
 Dim ErrDescr$
 If Fehlerzahl = 0 Then
  ErrDescr = Err.Description
  If oTask Is Nothing Then
   rufauf App.Path & "\nachricht.exe", "Fehler beim Planen von Tasks: " & ErrDescr, 0, , 0
   Exit Sub
  End If
'  Call Shell(App.Path & "\nachricht.exe " & "Fehler beim Einplanen von: " & oT.ApplicationName & vbCrLf & ErrDescr)
'  SuSh App.Path & "\nachricht.exe " & "Fehler beim Einplanen von: " & oT.ApplicationName & vbCrLf & ErrDescr, 0, , 0, 1
  rufauf App.Path & "\nachricht.exe", "Fehler beim Einplanen von: " & oTask.ApplicationName & vbCrLf & ErrDescr, 0, , 0
  Fehlerzahl = Fehlerzahl + 1
 End If
 Resume Next
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in xpmachTask/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' machTask

Function ListTasks(Sc As TaskScheduler2.Schedule)
#Const Excel = 0
 Dim MeldDatei$, MD2$, Cpt$, JName As SortierString, JListe As New SortierListe, machStr As New CString
 On Error GoTo fehler
 Cpt = CptName
 MeldDatei = uVerz & "ExistingTasks_" & Cpt & Format$(Now, "_dd.mm.yy_hh.mm.ss") & ".txt"
 MD2 = uVerz & "Task_Aufruf_" & Cpt & Format$(Now, "_dd.mm.yy_hh.mm.ss") & ".bas"
 Dim oT As Task, t As Trigger ' , erg&
 On Error Resume Next
#If Excel Then
 Close #332
 Open MeldDatei For Output As #332
 If Err.Number <> 0 Then Exit Function
#End If
 Close #334
 Open MD2 For Output As #334
 If Err.Number <> 0 Then Exit Function
 On Error GoTo fehler
#If Excel Then
 Print #332, "Name" & vbTab & "ApplicationName" & vbTab & "CommandLine" & vbTab & "Comment" & vbTab & "Flags" & vbTab & "MaxRunTime" & vbTab & "Priority" & vbTab & "WorkingDirectory" & vbTab & "Status"
 Print #332, vbTab & "Triggertyp" & vbTab & "BeginDay" & vbTab & "Daily/WeeklyInterval" & vbTab & "Weekly_DaysOfTheWeek" & vbTab & "Duration" & vbTab & "StartTime" & vbTab & "Text" & vbTab & "EndDay" & vbTab & "Flags" & vbTab & "Interval" & vbTab & "TriggerType"
#End If
 Print #334, "option explicit"
 Print #334, ""
 Print #334, "function Tasks_" & Cpt
 For Each oT In Sc
  Set JName = New SortierString
  JName.Stri = StrOhneUser(oT.name)
  If JListe.SuchItem(JName) <> -1 Then
   JListe.sCAdd JName
#If Excel Then
   Print #332, JName.Stri & vbTab & oT.ApplicationName & vbTab & oT.CommandLine & vbTab & oT.Comment & vbTab & oT.flags & vbTab & oT.MaxRunTime & vbTab & oT.Priority & vbTab & oT.WorkingDirectory & vbTab & oT.Status
#End If
   For Each t In oT.Triggers
'machAufgb(Sc As TaskScheduler2.Schedule, Name$, AName$, CdL$, flags&, MRT&, Prio&, WD$, Comment$, Creator$, BegD As Date, I1&, I2&, I3&, Itv&, Dur&, StartT As Date, EndD As Date, TFlags As TriggerFlags, TriggerType As TriggerTypes, Optional Stat As JobStatus = TaskReady)
    machStr.Clear
    machStr.AppVar Array("   machAufgb """, JName.Stri, """, """, oT.ApplicationName, """, """, REPLACE$(oT.CommandLine, """", """"""), """, ", oT.flags, ", ", oT.MaxRunTime, ", ", oT.Priority, ", """, oT.WorkingDirectory, """, """, REPLACE$(oT.Comment, """", """"""), """, """, REPLACE$(oT.Creator, """", """"""), """, ", Format(t.BeginDay, "\#mm\/dd\/yy\#"), ", ")
    Select Case t.TriggerType
     Case DailyTrigger:   machStr.AppVar Array(t.Daily_Interval, ",0 ,0 , ")
     Case WeeklyTrigger:  machStr.AppVar Array(t.Weekly_Interval, ", ", t.Weekly_DaysOfTheWeek, ",0 , ")
     Case MONTHLYDATE:    machStr.AppVar Array(t.MonthlyDate_Day, ", ", t.MonthlyDate_Months, ",0 , ")
     Case MonthlyWeekDay: machStr.AppVar Array(t.MonthlyDOW_DaysOfTheWeek, ", ", t.MonthlyDOW_Months, ", ", t.MonthlyDOW_Week, ", ")
     Case Else:           machStr.Append "0 ,0 ,0 , "
    End Select
    machStr.AppVar Array(t.Interval, ", ", t.Duration, ", ", Format(t.StartTime, "\#hh:mm:ss#"), ", ", Format(t.EndDay, "\#mm\/dd\/yy\#"), ", ", t.flags, ", ", t.TriggerType, ", ", oT.Status)
    Print #334, machStr
#If Excel Then
    Select Case t.TriggerType
     Case DailyTrigger
      Print #332, vbTab & "DailyTrigger:" & vbTab & t.BeginDay & vbTab & t.Daily_Interval & vbTab & "-" & vbTab & t.Duration & vbTab & t.StartTime & vbTab & t.Text & vbTab & t.EndDay & vbTab & t.flags & vbTab & t.Interval & vbTab & t.TriggerType
     Case WeeklyTrigger
      Print #332, vbTab & "WeeklyTrigger:" & vbTab & t.BeginDay & vbTab & t.Weekly_Interval & vbTab & t.Weekly_DaysOfTheWeek & vbTab & t.Duration & vbTab & t.StartTime & vbTab & t.Text & vbTab & t.EndDay & vbTab & t.flags & vbTab & t.Interval & vbTab & t.TriggerType
     Case Once
      Print #332, vbTab & "Once:" & vbTab & t.BeginDay & vbTab & "-" & vbTab & "-" & vbTab & t.Duration & vbTab & t.StartTime & vbTab & t.Text & vbTab & t.EndDay & vbTab & t.flags & vbTab & t.Interval & vbTab & t.TriggerType
     Case AtLogon
      Print #332, vbTab & "AtLogon:" & vbTab & t.BeginDay & vbTab & "-" & vbTab & "-" & vbTab & t.Duration & vbTab & t.StartTime & vbTab & t.Text & vbTab & t.EndDay & vbTab & t.flags & vbTab & t.Interval & vbTab & t.TriggerType
     Case t.TriggerType = AtSystemStart
      Print #332, vbTab & "AtSystemStart:" & vbTab & t.BeginDay & vbTab & "-" & vbTab & "-" & vbTab & t.Duration & vbTab & t.StartTime & vbTab & t.Text & vbTab & t.EndDay & vbTab & t.flags & vbTab & t.Interval & vbTab & t.TriggerType
     Case Else
      Print #332, "Sonstiger Trigger"
    End Select
#End If
   Next t
  End If
 Next oT
 Print #334, "End Function"
 Close #334
#If Excel Then
 Close #332
 Dim ExcelPfad$, cReg As New Registry
 ExcelPfad = cReg.ReadKey("Path", "SOFTWARE\Microsoft\Office\9.0\Excel\InstallRoot", HKEY_LOCAL_MACHINE)
' Call Shell(ExcelPfad & "Excel.exe " & MeldDatei, vbNormalFocus)
' SuSh ExcelPfad & "Excel.exe " & MeldDatei, 0, , 0
 rufauf ExcelPfad & "Excel.exe", MeldDatei
#End If
' Call Shell(Environ("systemroot") & "\system32\notepad.exe " & MD2, vbNormalFocus)
' SuSh Environ("systemroot") & "\system32\notepad.exe " & MD2, 0, , 0, 1
  rufauf Environ("systemroot") & "\system32\notepad", MD2, , , 0
 Exit Function
fehler:
ErrNumber = Err.Number
If ErrNumber = 429 Then
 On Error Resume Next
 Dim i%, FS$
 For i = 0 To 6
  Select Case i
   Case 0: FS = "\\linux1\daten\eigene Dateien\programmierung\tskschd.dll"
   Case 1: FS = "\\linmitte\programmierung\tskschd.dll"
   Case 2: FS = "\\linserv\programmierung\tskschd.dll"
   Case 3: FS = "\\mitte\u\programmierung\tskschd.dll"
   Case 4: FS = "\\anmeldl\U\programmierung\tskschd.dll"
   Case 5: FS = "\\anmeldr\eigene Dateien\programmierung\tskschd.dll"
   Case 6: FS = "\\anmeldrneu\u\programmierung\tskschd.dll"
  End Select
'  erg = fileexists(FS$)
'  If erg Then Exit For
  If FileExists(FS) Then Exit For
 Next i
 On Error GoTo 0
 If WV < win_vista Then
'  ShellaW "regsvr32.exe " & """" & FS & """", vbHide, , 100000
'  SuSh "regsvr32.exe " & """" & FS & """", 2, , 0
  rufauf "regsvr32", """" & FS & """", 2, , 0, 0
 Else
'  ShellaW doalsad & acceu & AdminGes & " cmd /e:on /c regsvr32 " & Chr$(34) & FS & Chr$(34), vbHide, , 100000
'  SuSh "cmd /e:on /c regsvr32 " & Chr$(34) & FS & Chr$(34), 1, , 0
  rufauf "cmd", "/e:on /c regsvr32 """ & FS & """", 2, , 0, 0
 End If
' WarteAufNicht "regsvr32", 100
' 27.8.15: Folgende Zeile könnte unnötig sein:
 schließ_direkt ("regsvr32")
 Resume
End If ' errnumber = 429
Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ListTasks/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ListTasks

Function Tasks()
#Const MySQLSichern = 0
#Const mitTMSichern = 0
#Const mitVerzVergl = 0
' Dim erg&
 Dim uName$
 On Error GoTo fehler
 uName = Environ("username")
 Call SetProgV
 If LenB(Cpt) = 0 Then Cpt = CptName
 Select Case Cpt
  Case "ANMELDL", "ANMELDL1", "NEUQUAD"
   anmeldlBoot = LWSuch("Boot")
   anmeldlData = LWSuch("Data")
   anmeldlVol = LWSuch("Data") ' LWSuch("Volume")
   If Date < #3/30/2012# Then
' TaskReady = 267008
    machAufgb "AutoFax", uVerz & "Programmierung\AutoFax\AutoFax.exe", "", 192, 259200000, 32, uVerz & "Programmierung\AutoFax", "", "schade", #8/10/2009#, 1, 0, 0, 10, 1493, #1:00:00 AM#, #11/30/1999#, 0, 1, 267008
'   machAufgb "BackupLösch", uVerz & "Programmierung\BackupLöschen\BackupLösch.exe", "", 192, 259200000, 32, uVerz & "Programmierung\BackupLöschen", "", "schade", #8/10/2009#, 1, 0, 0, 0, 0, #7:30:00 AM#, #11/30/1999#, 0, 1, 267008
    machAufgb "Faxakt ausführlich", uVerz & "Programmierung\FaxAkt\FaxAkt.exe", "", 192, 259200000, 32, uVerz & "Programmierung\FaxAkt", "", "schade", #8/11/2009#, 1, 0, 0, 1440, 1500, #12:01:00 AM#, #11/30/1999#, 0, 1, 267008
    machAufgb "Faxakt oft", uVerz & "Programmierung\FaxAkt\FaxAkt.exe", "nurneue", 192, 259200000, 32, uVerz & "Programmierung\FaxAkt", "", "schade", #8/12/2009#, 1, 0, 0, 5, 1015, #7:00:00 AM#, #11/30/1999#, 0, 1, 267008
    machAufgb "FaxDopp", uVerz & "Programmierung\FaxDopp\FaxDopp.exe", "mysql nachdb prot", 192, 259200000, 32, uVerz & "Programmierung\FaxAkt", "", "schade", #8/12/2009#, 1, 0, 0, 30, 960, #6:00:00 AM#, #11/30/1999#, 0, 1, 267008
    machAufgb "FaxDoppOrdner", uVerz & "Programmierung\FaxDopp\FaxDopp.exe mysql prot", "", 192, 259200000, 32, uVerz & "Programmierung\FaxDopp", "", uName, #9/24/2009#, 1, 0, 0, 60, 1440, #12:05:00 AM#, #11/30/1999#, 0, 1, 267009
   End If
#If MySQLSichern Then
    machAufgb "cluster", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpcluster"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxcluster""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:36:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "Desktop", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpDesktop"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxDesktop""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:37:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "dienstplan", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpdienstplan"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxdienstplan""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:23:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "dp", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpdp"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxdp""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #9/19/2007#, 1, 0, 0, 0, 0, #9:15:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "dpmitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpdpmitte"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxdpmitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 0, 0, 0, 0, #9:15:00 PM#, #11/30/1999#, 0, 1, 267010
    machAufgb "FaxeinP", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\Sturm.ANMELD2\Anwendungsdaten\MySQL\"" ""-clinux1praxis"" ""-bpfaxeinp"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxfaxeinp""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\Sturm.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/14/2008#, 1, 0, 0, 0, 0, #2:41:00 AM#, #11/30/1999#, 0, 1, 267008
    machAufgb "faxemitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpfaxemitte"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxfaxemitte""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 0, 0, 0, 0, #8:32:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "FotosinP", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\Sturm.ANMELD2\Anwendungsdaten\MySQL\"" ""-clinux1praxis"" ""-bpfotosinp"" ""-bt" & anmeldlVol & "\mysqlbackups\"" ""-bxfotosinp""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\Sturm.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #1/25/2008#, 1, 0, 0, 0, 0, #2:00:00 AM#, #11/30/1999#, 0, 1, 267008
    machAufgb "fotosinpmitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpfotosinpmitte"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxfotosinpmitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 0, 0, 0, 0, #8:39:00 PM#, #11/30/1999#, 0, 1, 267010
    machAufgb "gnr", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpgnr"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxgnr""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:34:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "gnrmitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpgnrmitte"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxgnrmitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 2, 0, 0, 0, #8:34:00 PM#, #11/30/1999#, 0, 2, 267010
    machAufgb "HAAkt", uVerz & "Programmierung\HAAkt\HAAkt.exe", "", 192, 259200000, 32, uVerz & "Programmierung\HAAkt", "", "schade", #7/6/2007#, 1, 0, 0, 0, 0, #11:00:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "icd10", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpicd10"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxicd10""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:35:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "icd10mitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpicd10mitte"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxicd10mitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 4095, 0, 0, 0, #3:47:00 AM#, #11/30/1999#, 0, 3, 267010
    machAufgb "KVAerzte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\Sturm.ANMELD2\Anwendungsdaten\MySQL\"" ""-clinux1praxis"" ""-bpkvaerzte"" ""-bt" & anmeldlVol & "\mysqlbackups\"" ""-bxkvaerzte""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\Sturm.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #1/25/2008#, 1, 0, 0, 0, 0, #1:30:00 AM#, #11/30/1999#, 0, 1, 267008
    machAufgb "kvaerztemitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpkvaerztemitte"" ""-bt" & anmeldlVol & "\mysqlbackups"" ""-bxkvaerztemitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 0, 0, 0, 0, #8:37:00 PM#, #11/30/1999#, 0, 1, 267010
    machAufgb "mysql", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpmysql"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxmysql""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:39:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "mysqlmitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpmysqlmitte"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxmysqlmitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 0, 0, 0, 0, #8:39:00 PM#, #11/30/1999#, 0, 1, 267010
    machAufgb "office", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpoffice"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxoffice""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:39:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "officeKompliziert", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpofficeKompliziert"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxofficeKompliziert""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:40:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "officemitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpofficemitte"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxofficemitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 0, 0, 0, 0, #8:39:00 PM#, #11/30/1999#, 0, 1, 267010
    machAufgb "programmstaende", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpprogrammstaende"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxprogrammstaende""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:40:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "programmstaendemitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpNew Project"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxprogrammstaendemitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 0, 0, 0, 0, #8:40:00 PM#, #11/30/1999#, 0, 1, 267010
    machAufgb "quelle", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\Sturm.ANMELD2\Anwendungsdaten\MySQL\"" ""-clinux1praxis"" ""-bpquelle"" ""-bt" & anmeldlVol & "\mysqlbackups"" ""-bxquelle""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\Sturm.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #1/25/2008#, 1, 0, 0, 0, 0, #11:00:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "quelle1", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpquelle1"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxquelle1""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:45:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "quelle1mitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpquelle1mitte"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxquelle1mitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 0, 0, 0, 0, #8:45:00 PM#, #11/30/1999#, 0, 1, 267010
    machAufgb "quelle2", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpquelle2"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxquelle2""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:53:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "quellemitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpquellemitte"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxquellemitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 0, 0, 0, 0, #8:41:00 PM#, #11/30/1999#, 0, 1, 267010
    machAufgb "statisches", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpstatisches"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxstatisches""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:59:00 PM#, #11/30/1999#, 0, 1, 267008
    machAufgb "statischesmitte", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpstatischesmitte"" ""-bt" & anmeldlVol & "\MySQLBackups"" ""-bxstatischesmitte""", 4, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELDl\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #12/30/2007#, 1, 4, 0, 0, 0, #8:59:00 PM#, #11/30/1999#, 0, 2, 267010
    machAufgb "SystemIdleDetector", ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application\RunAtSystemIdle.exe", "", 48, 259200000, 32, ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application", "", uName, #11/13/2008#, 0, 0, 0, 0, 0, #12:00:00 AM#, #11/30/1999#, 0, 5, 267008
    machAufgb "test", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bptest"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxtest""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #8:59:00 PM#, #11/30/1999#, 0, 1, 267008
#End If
#If mitTMSichern Then
   machAufgb "TMSichern auf GMX", uVerz & "Programmierung\TMSichern\TMSichern.exe", """m:\meine Dokumente\PraxisDB""", 196, 259200000, 32, uVerz & "Programmierung\TMSichern", "Turbomed auf GMX sichern", "gerald", #5/30/2006#, 0, 0, 0, 0, 0, #1:25:00 AM#, #11/30/1999#, 0, 0, 267010
   machAufgb "TMSichern", uVerz & "programmierung\tmsichern\TMSichern.exe", "" & anmeldlData & "\turbomed\praxisdb " & anmeldlData & "\turbomed\dokumente", 192, 259200000, 32, uVerz & "Programmierung\TMSichern", "", "schade", #6/25/2009#, 1, 0, 0, 240, 1185, #12:14:00 AM#, #11/30/1999#, 0, 1, 267008
'   machAufgb "TMSichern", uverz & "Programmierung\TMSichern\TMSichern.exe", """" & AnmeldlBoot & "\tmserv\PraxisDB"" """ & AnmeldlBoot & "\Tmserv\Dokumente""", 4100, -1, 32, uverz & "Programmierung\TMSichern", "", "gerald", #5/14/2006#, 1, 0, 0, 360, 1139, #1:00:00 AM#, #11/30/1999#, 0, 1, 267010
   machAufgb "TMWebSichern auf 150m", uVerz & "Programmierung\TMWebSichern\TMWebSichern.exe", """ftp.150m.com"" ""gschade.150m.com"" ""54AS13g"" ""V1""", 196, 259200000, 32, uVerz & "Programmierung\TMWebSichern", "", "gerald", #5/31/2006#, 1, 64, 0, 0, 0, #4:00:00 AM#, #11/30/1999#, 0, 2, 267010
   machAufgb "TMWebSichern auf LRZ", uVerz & "Programmierung\TMWebSichern\TMWebSichern.exe", """ftp.lrz-muenchen.de"" ""km601ao"" ""97a5o6"" ""webserver/webdata""", 4, 259200000, 32, uVerz & "Programmierung\TMSichern", "", "gerald", #5/31/2006#, 1, 0, 0, 0, 0, #3:36:00 AM#, #11/30/1999#, 0, 1, 267010
   machAufgb "TMWebSichern auf M-Net", uVerz & "Programmierung\TMWebSichern\TMWebSichern.exe", """home.mnet-online.de"" ""gschade"" ""17raga"" """"", 196, 259200000, 32, uVerz & "Programmierung\TMWebSichern", "", "gerald", #6/2/2006#, 0, 0, 0, 0, 0, #2:30:00 PM#, #11/30/1999#, 0, 0, 267010
   machAufgb "TMWebSichern auf my-place", uVerz & "Programmierung\TMWebSichern\TMWebSichern.exe", """ftp.my-place.us"" ""gschade"" ""54AS13g""", 4100, 259200000, 32, uVerz & "Programmierung\TMWebSichern", "", "gerald", #5/31/2006#, 1, 4, 0, 0, 0, #4:00:00 AM#, #11/30/1999#, 0, 2, 267010
   machAufgb "TMWebSichern Freenet", uVerz & "Programmierung\TMWebSichern\TMWebSichern.exe", """people-ftp.freenet.de"" ""gerald.schade"" ""97a5o6""", 196, 259200000, 32, uVerz & "Programmierung\TMSichern", "", "gerald", #6/2/2006#, 0, 0, 0, 0, 0, #4:00:00 AM#, #11/30/1999#, 0, 0, 267010
   machAufgb "TMWebSichern Lifeline", uVerz & "programmierung\TMWebsichern\TMWebSichern.exe", """ftp://gschade.liveline.de@gschade.liveline.de"" ""gschade.liveline.de"" ""54AS13g"" """"", 196, 259200000, 32, uVerz & "programmierung\TMWebsichern", "", "gerald", #6/5/2006#, 1, 8, 0, 0, 0, #4:00:00 AM#, #11/30/1999#, 0, 2, 267010
   machAufgb "TMWebSichern TerraHosting", uVerz & "Programmierung\TMWebSichern\TMWebSichern.exe", """user388.terra-hosting.de"" ""web388"" ""8MBhyeiQ"" ""files""", 196, 259200000, 32, uVerz & "Programmierung\TMWebSichern", "", "gerald", #6/2/2006#, 1, 40, 0, 0, 0, #4:36:00 AM#, #11/30/1999#, 0, 2, 267010
#End If
#Const mitTMbackup = 1
#If mitTMbackup Then
   machAufgb "turbomed", ProgVerz & "\MySQL\MySQL Tools for 5.0\MySQLAdministrator.exe", """-UD" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\"" ""-cLinux1"" ""-bpturbomed"" ""-bt" & anmeldlVol & "\MySQLBackups\"" ""-bxturbomed""", 0, -1, 32, "" & anmeldlBoot & "\Dokumente und Einstellungen\schade.ANMELD2\Anwendungsdaten\MySQL\", "", "MySQL Administrator", #7/15/2007#, 1, 0, 0, 0, 0, #9:01:00 PM#, #11/30/1999#, 0, 1, 267008
#End If
#If mitUserFeed Then
   machAufgb "User_Feed_Synchronization-{F25455A2-0F4F-43B4-9959-6C9155719508}", Environ("windir") & "\system32\msfeedssync.exe", "sync", 8704, 259200000, 32, ProgVerz & "\Internet Explorer", "Aktualisiert veraltete Systemfeeds.", uName, #8/10/2009#, 0, 0, 0, 5, 215, #8:25:00 PM#, #11/30/1999#, 0, 0, 267008
   machAufgb "User_Feed_Synchronization-{F25455A2-0F4F-43B4-9959-6C9155719508} 2", Environ("windir") & "\system32\msfeedssync.exe", "sync", 8704, 259200000, 32, ProgVerz & "\Internet Explorer", "Aktualisiert veraltete Systemfeeds.", uName, #8/11/2009#, 1, 0, 0, 5, 1440, #12:00:00 AM#, #11/30/1999#, 0, 1, 267008
#End If
#If mitVerzVergl Then
   machAufgb "Verzeichnisseangleichen LinuxSichern Oft", "u:\Programmierung\verzeichnissevergleichen\LSOftAnmeldL.bat", "", 0, 259200000, 32, uVerz & "Programmierung\verzeichnissevergleichen", "", "gerald", #6/4/2006#, 1, 0, 0, 0, 0, #2:00:00 PM#, #11/30/1999#, 0, 1, 267008 ' uVerz & "Programmierung\verzeichnissevergleichen\VerzeichnisseAngleichen.exe", ""LinuxSichern Oft AnmeldL.ini"" -l30
   machAufgb "Verzeichnisseangleichen LinuxSichern Oft 2", "u:\Programmierung\verzeichnissevergleichen\LSOft2AnmeldL.bat", "", 0, 259200000, 32, uVerz & "Programmierung\verzeichnissevergleichen", "", "gerald", #6/4/2006#, 1, 0, 0, 0, 0, #2:15:00 AM#, #11/30/1999#, 0, 1, 267008 '  uVerz & "Programmierung\verzeichnissevergleichen\VerzeichnisseAngleichen.exe", ""LinuxSichern Oft AnmeldL.ini"" -l30
'   machAufgb "VerzeichnisseAngleichen LinuxSichern Selten", "u:\Programmierung\verzeichnissevergleichen\LSSeltenAnmeldL.bat", """LinuxSichern Selten AnmeldL.ini"" -l60", 192, 259200000, 32, uVerz & "Programmierung\VerzeichnisseVergleichen", "", "gerald", #6/4/2006#, 1, 64, 0, 0, 0, #3:30:00 AM#, #11/30/1999#, 0, 2, 267008 ' uVerz & "Programmierung\VerzeichnisseVergleichen\VerzeichnisseAngleichen.exe"
#End If
#If mitZuMailen Then
   machAufgb "zumailen", uVerz & "programmierung\zumailen\ZuMailen.exe", "", 192, 259200000, 32, uVerz & "programmierung\zumailen", "", uName, #7/25/2008#, 1, 0, 0, 10, 1439, #12:06:00 AM#, #7/25/2008#, 1, 1, 267008
#End If
  Case "ANMELDR"
   anmeldrZweit = LWSuch("Zweit")
   anmeldrBackup = LWSuch("BACKUP")
   machAufgb "AppleSoftwareUpdate", anmeldrBackup & "\Programme\Apple Software Update\SoftwareUpdate.exe", "-task", 4, 259200000, 32, "", "", "SYSTEM", #9/9/2008#, 1, 8, 0, 0, 0, #2:02:00 PM#, #11/30/1999#, 0, 2, 267010
   machAufgb "SystemIdleDetector", ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application\RunAtSystemIdle.exe", "", 48, 259200000, 32, ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application", "", "gschade", #7/8/2008#, 0, 0, 0, 0, 0, #12:00:00 AM#, #11/30/1999#, 0, 5, 267008
#If mitTMSichern Then
   machAufgb "TMSichern oft", uVerz & "Programmierung\TMSichern\TMSichern.exe", anmeldrZweit & "\turbomed\PraxisDB", 192, 259200000, 32, uVerz & "Programmierung\TMSichern", "", uName, #2/21/2009#, 1, 0, 0, 240, 1412, #9:27:00 AM#, #11/30/1999#, 0, 1, 267008
   machAufgb "TMSichern selten", uVerz & "Programmierung\TMSichern\TMSichern.exe", anmeldrZweit & "\turbomed\praxisdb " & anmeldrZweit & "\turbomed\dokumente", 192, 259200000, 32, uVerz & "Programmierung\TMSichern", "", uName, #2/21/2009#, 1, 0, 0, 0, 0, #5:00:00 AM#, #11/30/1999#, 0, 1, 267008
#End If
  Case "ANMELDRNEU"
   anmeldrZweit = LWSuch("DATA")
   anmeldrBackup = LWSuch("DATA")
   machAufgb "Sa u So ausschalten", Environ("windir") & "\system32\shutdown.exe", " -s -t 01", 4, 259200000, 32, "", "", "SYSTEM", #9/9/2008#, 1, 8, 0, 0, 0, #2:02:00 PM#, #11/30/1999#, 0, 2, 267010
   machAufgb "AppleSoftwareUpdate", anmeldrBackup & "\Programme\Apple Software Update\SoftwareUpdate.exe", "-task", 4, 259200000, 32, "", "", "SYSTEM", #9/9/2008#, 1, 8, 0, 0, 0, #2:02:00 PM#, #11/30/1999#, 0, 2, 267010
   machAufgb "SystemIdleDetector", ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application\RunAtSystemIdle.exe", "", 48, 259200000, 32, ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application", "", "gschade", #7/8/2008#, 0, 0, 0, 0, 0, #12:00:00 AM#, #11/30/1999#, 0, 5, 267008
#If mitTMSichern Then
   machAufgb "TMSichern oft", uVerz & "Programmierung\TMSichern\TMSichern.exe", anmeldrZweit & "\turbomed\PraxisDB", 192, 259200000, 32, uVerz & "Programmierung\TMSichern", "", uName, #2/21/2009#, 1, 0, 0, 240, 1412, #9:27:00 AM#, #11/30/1999#, 0, 1, 267008
   machAufgb "TMSichern selten", uVerz & "Programmierung\TMSichern\TMSichern.exe", anmeldrZweit & "\turbomed\praxisdb " & anmeldrZweit & "\turbomed\dokumente", 192, 259200000, 32, uVerz & "Programmierung\TMSichern", "", uName, #2/21/2009#, 1, 0, 0, 0, 0, #5:00:00 AM#, #11/30/1999#, 0, 1, 267008
#End If
  Case "ANMELDR1"
   anmeldrZweit = LWSuch("Data") ' D:
   anmeldrBackup = LWSuch("Data") ' D:
   machAufgb "AppleSoftwareUpdate", anmeldrBackup & "\Programme\Apple Software Update\SoftwareUpdate.exe", "-task", 4, 259200000, 32, "", "", "SYSTEM", #9/9/2008#, 1, 8, 0, 0, 0, #2:02:00 PM#, #11/30/1999#, 0, 2, 267010
   machAufgb "SystemIdleDetector", ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application\RunAtSystemIdle.exe", "", 48, 259200000, 32, ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application", "", "gschade", #7/8/2008#, 0, 0, 0, 0, 0, #12:00:00 AM#, #11/30/1999#, 0, 5, 267008
#If mitTMSichern Then
   machAufgb "TMSichern oft", uVerz & "Programmierung\TMSichern\TMSichern.exe", anmeldrZweit & "\turbomed\PraxisDB", 192, 259200000, 32, uVerz & "Programmierung\TMSichern", "", uName, #2/21/2009#, 1, 0, 0, 240, 1412, #9:27:00 AM#, #11/30/1999#, 0, 1, 267008
   machAufgb "TMSichern selten", uVerz & "Programmierung\TMSichern\TMSichern.exe", anmeldrZweit & "\turbomed\praxisdb " & anmeldrZweit & "\turbomed\dokumente", 192, 259200000, 32, uVerz & "Programmierung\TMSichern", "", uName, #2/21/2009#, 1, 0, 0, 0, 0, #5:00:00 AM#, #11/30/1999#, 0, 1, 267008
#End If
#If alteMitte Then
  Case "MITTE", "MITTE1"
   mitteVol = LWSuch("Volume")
   mitteRoot = LWSuch("Root")
   machAufgb "AppleSoftwareUpdate", ProgVerz & "\Apple Software Update\SoftwareUpdate.exe", "-task", 4, 259200000, 32, "", "", "SYSTEM", #8/14/2008#, 1, 16, 0, 0, 0, #11:27:00 AM#, #11/30/1999#, 0, 2, TaskDisabled
   machAufgb "BDTkompr", uVerz & "Programmierung\BDTkompr\BDTkompr.exe", "/auto -u- -o- -e BDT -v u:\tmexport -d 5", 192, 259200000, 32, uVerz & "Programmierung\BDTkompr", "", "schade", #8/14/2009#, 1, 0, 0, 0, 0, #12:14:00 AM#, #11/30/1999#, 0, 1, 267011
   machAufgb "Google Software Updater", ProgVerz & "\Google\Common\Google Updater\GoogleUpdaterService.exe", "scheduled_start", 0, -1, 32, "", "Mit Google Updater bleibt Ihre Google-Software stets auf dem neuesten Stand. Wird der Google Updater-Service deaktiviert oder angehalten, so wird Ihre Google-Software nicht mehr aktualisiert, was dazu führen kann, dass etwaige Sicherheitslücken nicht geschlossen werden und bestimmte Funktionen möglicherweise nicht mehr verfügbar sind.", "SYSTEM", #8/16/2009#, 1, 0, 0, 0, 144000, #1:01:00 PM#, #11/30/1999#, 0, 1, 267011
   machAufgb "Google Software Updater", ProgVerz & "\Google\Common\Google Updater\GoogleUpdaterService.exe", "scheduled_start", 0, -1, 32, "", "Mit Google Updater bleibt Ihre Google-Software stets auf dem neuesten Stand. Wird der Google Updater-Service deaktiviert oder angehalten, so wird Ihre Google-Software nicht mehr aktualisiert, was dazu führen kann, dass etwaige Sicherheitslücken nicht geschlossen werden und bestimmte Funktionen möglicherweise nicht mehr verfügbar sind.", "SYSTEM", #8/17/2009#, 0, 0, 0, 20, 144000, #7:23:00 PM#, #11/30/1999#, 0, 0, 267011
   machAufgb "GoogleUpdateTaskMachineCore", ProgVerz & "\Google\Update\GoogleUpdate.exe", "/c", 0, -1, 32, "", "Hält Ihre Google-Software auf dem neuesten Stand. Wenn diese Anwendung deaktiviert oder angehalten wird, wird Ihre Google-Software nicht aktualisiert. Das heißt, dass eventuell auftretende Sicherheitslücken nicht behoben und bestimmte Funktionen möglicherweise nicht ausgeführt werden können. Diese Anwendung deinstalliert sich selbst, wenn sie nicht von einer Google-Software verwendet wird.", "SYSTEM", #1/1/1999#, 0, 0, 0, 0, 0, #12:00:00 AM#, #11/30/1999#, 0, 7, 267008
   machAufgb "GoogleUpdateTaskMachineCore", ProgVerz & "\Google\Update\GoogleUpdate.exe", "/c", 0, -1, 32, "", "Hält Ihre Google-Software auf dem neuesten Stand. Wenn diese Anwendung deaktiviert oder angehalten wird, wird Ihre Google-Software nicht aktualisiert. Das heißt, dass eventuell auftretende Sicherheitslücken nicht behoben und bestimmte Funktionen möglicherweise nicht ausgeführt werden können. Diese Anwendung deinstalliert sich selbst, wenn sie nicht von einer Google-Software verwendet wird.", "SYSTEM", #8/3/2009#, 1, 0, 0, 0, 0, #11:32:00 PM#, #11/30/1999#, 0, 1, 267008
   machAufgb "GoogleUpdateTaskMachineUA", ProgVerz & "\Google\Update\GoogleUpdate.exe", "/ua /installsource scheduler", 0, -1, 32, "", "Hält Ihre Google-Software auf dem neuesten Stand. Wenn diese Anwendung deaktiviert oder angehalten wird, wird Ihre Google-Software nicht aktualisiert. Das heißt, dass eventuell auftretende Sicherheitslücken nicht behoben und bestimmte Funktionen möglicherweise nicht ausgeführt werden können. Diese Anwendung deinstalliert sich selbst, wenn sie nicht von einer Google-Software verwendet wird.", "SYSTEM", #8/3/2009#, 1, 0, 0, 60, 1440, #11:32:00 PM#, #11/30/1999#, 0, 1, 267008
   machAufgb "GoogleUpdateTaskUserS-1-5-21-1801674531-602609370-725345543-1011Core", mitteRoot & "\Dokumente und Einstellungen\gerald\Lokale Einstellungen\Anwendungsdaten\Google\Update\GoogleUpdate.exe", "/c", 8192, -1, 32, "", "Hält Ihre Google-Software auf dem neuesten Stand. Wenn diese Anwendung deaktiviert oder angehalten wird, wird Ihre Google-Software nicht aktualisiert. Das heißt, dass eventuell auftretende Sicherheitslücken nicht behoben und bestimmte Funktionen möglicherweise nicht ausgeführt werden können. Diese Anwendung deinstalliert sich selbst, wenn sie nicht von einer Google-Software verwendet wird.", "gerald", #7/1/2009#, 1, 0, 0, 0, 0, #3:50:00 AM#, #11/30/1999#, 0, 1, 267008
   machAufgb "GoogleUpdateTaskUserS-1-5-21-1801674531-602609370-725345543-1011UA", mitteRoot & "C\Dokumente und Einstellungen\gerald\Lokale Einstellungen\Anwendungsdaten\Google\Update\GoogleUpdate.exe", "/ua /installsource scheduler", 8192, -1, 32, "", "Hält Ihre Google-Software auf dem neuesten Stand. Wenn diese Anwendung deaktiviert oder angehalten wird, wird Ihre Google-Software nicht aktualisiert. Das heißt, dass eventuell auftretende Sicherheitslücken nicht behoben und bestimmte Funktionen möglicherweise nicht ausgeführt werden können. Diese Anwendung deinstalliert sich selbst, wenn sie nicht von einer Google-Software verwendet wird.", "gerald", #7/1/2009#, 1, 0, 0, 60, 1440, #3:50:00 AM#, #11/30/1999#, 0, 1, 267008
   machAufgb "HDI alt komprimieren mit BDTKompr", uVerz & "Programmierung\BDTkompr\BDTkompr.exe", "-u -o- -d 30 -v p:\hdi alt -e * -auto", 192, 259200000, 32, uVerz & "Programmierung\BDTkompr", "", "schade", #8/14/2009#, 1, 0, 0, 0, 0, #2:30:00 AM#, #11/30/1999#, 0, 1, 267011
   machAufgb "PBereinigen i.dat.p", uVerz & "Programmierung\PBereinigen\PBereinigen.exe", "-p " & mitteVol & "\dat\p -pe " & mitteVol & "\dat\p\eingelesen -tm " & mitteVol & "\turbomed -hd " & mitteVol & "\dat\p\hdi alt -tx " & mitteVol & "\dat\u\tmexport -auto", 192, 259200000, 32, uVerz & "Programmierung\PBereinigen", "", "schade", #8/14/2009#, 1, 0, 0, 0, 0, #4:00:00 PM#, #11/30/1999#, 0, 1, 267008
   machAufgb "PBereinigen p.", uVerz & "Programmierung\PBereinigen\PBereinigen.exe", "-p p:\ -pe p:\eingelesen -tm \\linux1\turbomed -hd p:\hdi alt -tx u:\tmexport -auto", 192, 259200000, 32, uVerz & "Programmierung\PBereinigen", "", "schade", #8/14/2009#, 1, 0, 0, 0, 0, #9:00:00 PM#, #11/30/1999#, 0, 1, 267008
   machAufgb "SystemIdleDetector", ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application\RunAtSystemIdle.exe", "", 48, 259200000, 32, ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application", "", "schade", #7/5/2008#, 0, 0, 0, 0, 0, #12:00:00 AM#, #11/30/1999#, 0, 5, 267008
#End If
#If mitTMSichern Then
   machAufgb "TMSichern", uVerz & "programmierung\tmsichern\tmsichern.exe", mitteVol & "\turbomed\praxisdb " & mitteVol & "\dat\turbomed\dokumente", 192, 259200000, 32, uVerz & "programmierung\tmsichern", "", "schade", #8/14/2009#, 1, 0, 0, 0, 0, #12:14:00 AM#, #11/30/1999#, 0, 1, 267011
#End If
   machAufgb "User_Feed_Synchronization-{1F363954-58AA-47E3-8722-209B4ED43FBE}", Environ("windir") & "\system32\msfeedssync.exe", "sync", 8704, 259200000, 32, ProgVerz & "\Internet Explorer", "Aktualisiert veraltete Systemfeeds.", uName, #8/17/2009#, 0, 0, 0, 5, 1208, #3:52:00 AM#, #11/30/1999#, 0, 0, 267008
   machAufgb "User_Feed_Synchronization-{1F363954-58AA-47E3-8722-209B4ED43FBE}", Environ("windir") & "\system32\msfeedssync.exe", "sync", 8704, 259200000, 32, ProgVerz & "\Internet Explorer", "Aktualisiert veraltete Systemfeeds.", uName, #8/18/2009#, 1, 0, 0, 5, 1440, #12:03:00 AM#, #11/30/1999#, 0, 1, 267008
#If mitVerzVergl Then
   machAufgb "VerzAngl Oft", "u:\Programmierung\verzeichnissevergleichen\LSOftMitte.bat", "", 192, 259200000, 32, uVerz & "Programmierung\verzeichnissevergleichen", "", "schade", #8/14/2009#, 1, 0, 0, 0, 0, #2:15:00 PM#, #11/30/1999#, 0, 1, 267011 ' uVerz & "Programmierung\verzeichnissevergleichen\VerzeichnisseAngleichen.exe", """Linuxsichern oft Mitte.ini"""
   machAufgb "VerzAngl Selten", "u:\Programmierung\verzeichnissevergleichen\LSSeltenMitte.bat", "", 192, 259200000, 32, "" & mitteVol & "\DAT\U\Programmierung\verzeichnissevergleichen", "", "schade", #8/14/2009#, 1, 0, 0, 0, 0, #6:35:00 AM#, #11/30/1999#, 0, 1, 267008 ' "" & mitteVol & "\DAT\U\Programmierung\verzeichnissevergleichen\VerzeichnisseAngleichen.exe", """LinuxSichern Selten Mitte.ini"""
#End If
  Case "SONO"
   sonoBoot = LWSuch("Boot")
   sonoDaten = LWSuch("Daten")
   machAufgb "cmd", Environ("windir") & "\system32\cmd.exe", "/c del /s /f /q \\linux1\daten\papierkorb\*.ldb", 192, 259200000, 32, Environ("windir") & "\system32", "*.ldb-Dateien löschen", uName, #5/31/2009#, 1, 0, 0, 0, 0, #2:30:00 PM#, #11/30/1999#, 0, 1, 267008
   machAufgb "PBereinigen", uVerz & "Programmierung\PBereinigen\PBereinigen.exe", "-p " & sonoDaten & "\P\ -pe " & sonoDaten & "\p\eingelesen -tm " & sonoDaten & "\turbomed -hd " & sonoDaten & "\p\hdi alt -tx " & sonoDaten & "\u\tmexport /auto", 192, 259200000, 32, uVerz & "Programmierung\PBereinigen", "", uName, #5/31/2009#, 1, 0, 0, 0, 0, #3:00:00 PM#, #11/30/1999#, 0, 1, 267008
   machAufgb "SystemIdleDetector", ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application\RunAtSystemIdle.exe", "", 48, 259200000, 32, ProgVerz & "\Roche Diagnostics\ACCU-CHEK 360\Application", "", uName, #7/7/2008#, 0, 0, 0, 0, 0, #12:00:00 AM#, #11/30/1999#, 0, 5, 267008
#If mitVerzVergl Then
   machAufgb "VerzeichnisseAngleichen", uVerz & "Programmierung\VerzeichnisseVergleichen\VerzeichnisseAngleichen.exe", "sono.ini", 196, 259200000, 32, uVerz & "Programmierung\VerzeichnisseVergleichen", "", "Schade", #11/9/2007#, 1, 0, 0, 0, 0, #2:00:00 PM#, #11/30/1999#, 0, 1, 267010
#End If
 End Select
 KWn "poetaktiv.exe", uVerz & "Programmierung\poetaktiv", ProgVerz
 KWn "Expaufruf.exe", uVerz & "Programmierung\ExpAufruf", ProgVerz
 machAufgb "Turbomed Ausfallwarnung", ProgVerz & "poetaktiv.exe", "", 8193, 1800000, 32, ProgVerz, "", uName, #9/2/2015#, 1, 0, 0, 2, 1439, #12:01:00 AM#, #11/30/1999#, 0, 1, 267008, True, "", False
' rufauf "cmd", "/c schtasks /query /tn ""Turbomed toeten"" >NUL 2>&1 || schtasks /create /ru administrator /rp " & AdminPwd & " /sc minute /tn ""Turbomed toeten"" /tr ""cmd /c 'if not exist \\virtwin\turbomed\lauf taskkill /im turbomed.exe /t /f'""", True
 machAufgb "Turbomed töten", "cmd", "", 8193, 1800000, 1, "", "Turbomed für Backup beenden", uName, #10/13/2021#, 1, 0, 0, 1, 1439, #12:02:00 AM#, #11/30/1999#, 0, 1, 267008, True, "/c ""if exist \\linux1\turbomed\lau findstr /c:Mehrplatzbetrieb={ja} c:\turbomed\programm\local.ini >Nul 2>&amp;1 &amp;&amp; taskkill /im turbomed.exe /t /f""", True, True
 Exit Function
 ' On Error Resume Next
 'machAufgb "Turbomed Ausfallwarnung", ProgVerz & "\poetaktiv.exe", "", 0, 259200000, 32, ProgVerz, vNS, uname, #6/24/2010#, 1, 0, 0, 2, 1439, #12:00:00 AM#, #11/30/1999#, 0, 1, 267011 ' "debug", am 10.11.11 entfernt
fehler:
ErrNumber = Err.Number
If ErrNumber = 429 Then
 On Error Resume Next
 Dim i%, FS$
 For i = 0 To 6
  Select Case i
   Case 0: FS = "\\linux1\daten\eigene Dateien\programmierung\tskschd.dll"
   Case 1: FS = "\\linmitte\programmierung\tskschd.dll"
   Case 2: FS = "\\linserv\programmierung\tskschd.dll"
   Case 3: FS = "\\mitte\u\programmierung\tskschd.dll"
   Case 4: FS = "\\anmeldl\U\programmierung\tskschd.dll"
   Case 5: FS = "\\anmeldr\eigene DAteien\programmierung\tskschd.dll"
   Case 6: FS = "\\anmeldrneu\u\programmierung\tskschd.dll"
   Case 7: FS = "\\anmeldr1\u\programmierung\tskschd.dll"
  End Select
'  erg = fileexists(FS$)
'  If erg Then Exit For
  If FileExists(FS) Then Exit For
 Next i
 On Error GoTo 0
 If WV < win_vista Then
'  ShellaW "regsvr32.exe " & """" & FS & """", vbHide, , 100000
'  SuSh "regsvr32.exe " & """" & FS & """", , , 0
  rufauf "regsvr32", """" & FS & """", , , 0, 0
 Else
'  ShellaW doalsad & acceu & AdminGes & " cmd /e:on /c regsvr32 " & Chr$(34) & FS & Chr$(34), vbHide, , 100000
'  SuSh "cmd /e:on /c regsvr32 " & Chr$(34) & FS & Chr$(34), 1, , 0
   rufauf "cmd", "/e:on /c regsvr32 """ & FS & """", 2, , 0, 0
 End If
' WarteAufNicht "regsvr32", 100
' 27.8.15: folgende Zeile könnte wieder unnötig sein:
 schließ_direkt ("regsvr32")
 Resume
End If
Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Tasks/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Tasks

Sub SetProgV2()
 On Error GoTo fehler
 AUP = Environ("allusersprofile")
 Dim APD$
 APD = Environ("appdata")
 userprof = Environ("userprofile")
 If WV < win_vista Then
  StartMen = AUP & "\Startmenü"
  StartMenProg = StartMen & "\Programme"
  autoVz = StartMenProg & "\Autostart"
  Favor = userprof & "\Favoriten\"
 Else
  StartMen = AUP & "\Microsoft\Windows\Start Menu"
  StartMenProg = StartMen & "\Programs"
'  autoVz = StartMenProg & "\Startup"
  autoVz = APD & "\Microsoft\Windows\Start Menu\Programs\Startup"
  Favor = userprof & "\Favorites\"
#Const Edge = False
#If Edge Then
  Dim EdgeFav$
  EdgeFav = Environ("LOCALAPPDATA") + "\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\User\Default\Favorites"
  If FSO.FolderExists(EdgeFav) Then
   Favor = EdgeFav
  Else
#End If
   Favor = GetSpecialFolder(ssfFAVORITES)
#If Edge Then
  End If
#End If
  VerzPrüf (Favor)
 End If
 UN = UserName
 FPos = 2
 Cpt = CptName
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SetProgV2/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' SetProgV2()

Function SetProgV3&()
 On Error GoTo fehler
 LWTrekstor = LWSuch("TREKSTOR")
 Select Case Cpt
  Case "ANMELDR1"
   arBackup = LWSuch("Rest") ' E:
   arRecover = LWSuch("Root") ' C:
   arDasi = arRecover & "\Turbomed-DASI"
   arDokumente = arBackup & "\turbomed\Dokumente" ' h:\turbomed\dokumente
   arData = LWSuch("Root")  ' C: 'arBoot = LWSuch("BOOT")
  Case "ANMELDR"
   arBackup = LWSuch("BACKUP") ' H:
   arRecover = LWSuch("RECOVER") ' E:
   arDasi = arRecover & "\Turbomed-DASI"
   arDokumente = arBackup & "\turbomed\Dokumente" ' h:\turbomed\dokumente
   arData = LWSuch("DATA")  'arBoot = LWSuch("BOOT")
  Case "ANMELDRNEU"
   arBackup = LWSuch("DATA") ' E:
   arRecover = LWSuch("DATA") ' E:
   arDasi = arRecover & "\Turbomed-DASI"
   arDokumente = arBackup & "\turbomed\Dokumente" ' h:\turbomed\dokumente
   arData = LWSuch("DATA")
  Case "MITTE"
   mitteVol = LWSuch("Volume")
   mitteRoot = LWSuch("Root")
   mitteAustausch = LWSuch("Austausch")
  Case "MITTE1"
   mitteVol = LWSuch("Boot")
   mitteRoot = LWSuch("Boot")
   mitteAustausch = LWSuch("Boot")
  Case "ANMELDL", "ANMELDL1"
   alBoot = LWSuch("Boot")
   alData = LWSuch("Data")
   alDasi = alData & "\TM-DASI"
   alVolume = LWSuch("Boot") ' LWSuch("Volume")
   EigDatAnmL = alData & "\daten\eigene Dateien" ' alVolume & "\eigene Dateien alt"
   DatenAnmL = alData & "\daten"
   PatDokAnmL = alData & "\daten\patientendokumente" ' alVolume & "\P"
   FaxeZwischen = alVolume & "\Faxe Zwischen"
   DownAnmL = alVolume & "\down" ' alData & "\daten\down"
   GeraldAnmL$ = alData & "\daten\shome\gerald"
   KothnyAnmL$ = alData & "\daten\shome\kothny"
   ReadOAnmL$ = alData & "\daten\shome\gerald"
   ReadOKAnmL$ = alData & "\daten\shome\kothny"
 End Select
 
 If (FileExists("\\linux1\obsläuft\Läuft") Or FileExists("\\linux1\obslaeuft\Laeuft")) Then
  obNot = 0
'  If Not IsIDE Then If InStr(LCase(App.Path), "\\linux1") = 0 Then Exit Sub
  TMServCpt = "Linux1"
  TMStammVk = "\turbomed"
  TMStammV = "\\" & TMServCpt & TMStammVk ' "\\Linux1\Turbomed"
  PatDok = "\\linux1\daten\Patientendokumente" '"\\linux1\Gemein\patdok" '/ "\\MITTE\p"
  EigDat = "\\linux1\daten\eigene Dateien" '"\\linux1\Gemein\eigene Dateien" ' "\\MITTE\u"
  Down = "\\linux1\daten\down" ' "\\linux1\Gemein\down" ' "\\MITTE\v"
  Gerald = "\\linux1\daten\shome\gerald" '\\ANMELDR\Gerald
  Kothny = "\\linux1\daten\shome\kothny" '\\ANMELDr\Kothny
  ReadO = "\\linux1\geraldprivat" '\\ANMELDR\Gerald
  ReadOK = "\\linux1\kothnyprivat" '\\ANMELDR\Gerald
  Programme = "\\linux1\Daten\Programme" '"\\linux1\Gemein\Programme" '"\\MITTE\Gemein\Programme")
  Dokumente = "\\linux1\daten\turbomed\Dokumente" '"\\linux1\Gemein\Dokumente"
  Call fgetgsreg("LINUX1")
 ElseIf FileExists("\\linmitte\obsläuft\Läuft") Then
  obNot = 1
  TMServCpt = "Linmitte"
  TMStammVk = "\turbomed"
  TMStammV = "\\" & TMServCpt & TMStammVk ' "\\Linux1\Turbomed"
  PatDok = "\\linmitte\sam\p"
  EigDat = "\\linmitte\sam\eigene Dateien"
  Down = "\\linmitte\sam\down"
  Gerald = "\\linmitte\sam\gerald" '\\ANMELDR\Gerald
  Kothny = "\\linmitte\sam\kothny" '\\ANMELDr\Kothny
  ReadO = "\\linmitte\geraldprivat" '\\ANMELDR\Gerald
  ReadOK = "\\linmitte\kothnyprivat" '\\ANMELDR\Gerald
  Programme = "\\linmitte\sam\programme" '"\\linux1\Gemein\Programme" '"\\mitte\Gemein\Programme")
  Dokumente = "\\linmitte\daten\turbomed\Dokumente" '"\\linux1\Gemein\Dokumente"
'  Shell (App.Path + "\Nachricht.exe")
'  SuSh App.Path + "\Nachricht.exe", , , 0, 1
  rufauf App.Path & "\Nachricht.exe", , , , 0
  Call fgetgsreg("LINMITTE")
 ElseIf FileExists("\\linserv\obsläuft\Läuft") Then
  obNot = 2
  TMServCpt = "linserv"
  TMStammVk = "\turbomed"
  TMStammV = "\\" & TMServCpt & TMStammVk ' "\\Linux1\Turbomed"
  PatDok = "\\linserv\daten\p"
  EigDat = "\\linserv\daten\eigene Dateien"
  Down = "\\linserv\daten\down"
  Gerald = "\\linserv\daten\gerald" '\\ANMELDR\Gerald
  Kothny = "\\linserv\daten\kothny"
  ReadO = "\\linserv\geraldprivat" '\\ANMELDR\Gerald
  ReadOK = "\\linserv\kothnyprivat"
  Programme = "\\linserv\daten\programme" '"\\linux1\Gemein\Programme" '"\\MITTE\Gemein\Programme")
  Dokumente = "\\linserv\daten\turbomed\Dokumente" '"\\linux1\Gemein\Dokumente"
'  Shell (App.Path + "\Nachricht.exe")
'  SuSh App.Path + "\Nachricht.exe", , , 0, 1
  rufauf App.Path & "\Nachricht.exe", , , , 0
  Call fgetgsreg("LINSERV")
 ElseIf FileExists("\\MITTE\obsläuft\läuft") Then
  obNot = -2
  If Not IsIDE Then
   If InStr(LCase(App.Path), "\\linux1") <> 0 Then
    GoTo schluss
   End If
  End If
  PatDok = "\\MITTE\P"
  EigDat = "\\MITTE\U"
  Down = "\\MITTE\v"
  Gerald = "\\MITTE\Gerald"
  Kothny = "\\MITTE\Kothny"
  ReadO = "\\MITTE\geraldprivat"
  ReadOK = "\\MITTE\kothnyprivat"
  Programme = "\\MITTE\Gemein\Programme"
  If FileExists("\\MITTE\Turbomed\PraxisDB\objects.dat") Then
   TMServCpt = "MITTE"
   TMStammVk = "\Turbomed"
   TMStammV = "\\" & TMServCpt & TMStammVk ' "\\MITTE\Turbomed"
   Dokumente = "\\MITTE\Turbomed\Dokumente"
  End If
'  Call Shell(TMStammV & "\Programm\" & "ptserv32.exe")
'  SuSh TMStammV & "\Programm\" & "ptserv32.exe", , , 0
   rufauf TMStammV & "\Programm\" & "ptserv32.exe", , , , 0, 0
'  Shell (App.Path + "\Nachricht.exe")
'  SuSh App.Path + "\Nachricht.exe", , , 0, 1
  rufauf App.Path & "\Nachricht.exe", , , , 0
  Call fgetgsreg("MITTE")
 ElseIf FileExists("\\MITTE1\obsläuft\läuft") Then
  obNot = -2
  If Not IsIDE Then
   If InStr(LCase(App.Path), "\\linux1") <> 0 Then
    GoTo schluss
   End If
  End If
  PatDok = "\\MITTE1\P"
  EigDat = "\\MITTE1\U"
  Down = "\\MITTE1\v"
  Gerald = "\\MITTE1\Gerald"
  Kothny = "\\MITTE1\Kothny"
  ReadO = "\\MITTE1\geraldprivat"
  ReadOK = "\\MITTE1\kothnyprivat"
  Programme = "\\MITTE1\Gemein\Programme"
  If FileExists("\\MITTE1\Turbomed\PraxisDB\objects.dat") Then
   TMServCpt = "MITTE1"
   TMStammVk = "\Turbomed"
   TMStammV = "\\" & TMServCpt & TMStammVk ' "\\MITTE\Turbomed"
   Dokumente = "\\MITTE1\Turbomed\Dokumente"
  End If
'  SuSh TMStammV & "\Programm\" & "ptserv32.exe", , , 0
   rufauf TMStammV & "\Programm\" & "ptserv32.exe", , , , 0, 0
'  Shell (App.Path + "\Nachricht.exe")
'  SuSh App.Path + "\Nachricht.exe", , , 0, 1
  rufauf App.Path & "\Nachricht.exe", , , , 0
  Call fgetgsreg("MITTE1")
 Else
  obNot = -1
  If Not IsIDE Then If InStr(LCase(App.Path), "\\linux1") <> 0 Then GoTo schluss ' falls Linux1 existiert aber nicht Linux1\Gemein ...
  PatDok = "\\ANMELDL\P"
  EigDat = "\\ANMELDL\U"
  Down = "\\ANMELDL\down"
  Gerald = "\\ANMELDL\Gerald"
  Kothny = "\\ANMLEDL\Kothny"
  ReadO = "\\ANMELDL\geraldprivat"
  ReadOK = "\\ANMELDL\kothnyprivat"
  Programme = "\\ANMELDL\Gemein\Programme"
  If FileExists("\\ANMELDRNEU\turbomed\programm\turbomed.exe") Then
   TMServCpt = "ANMELDRNEU"
   TMStammVk = "\turbomed"
   TMStammV = "\\" & TMServCpt & TMStammVk ' "\\Linux1\Turbomed"
   Dokumente = "\\ANMELDRNEU\Turbomed\Dokumente"
   Call fgetgsreg("ANMELDRNEU")
  ElseIf FileExists("\\ANMELDR\turbomed\programm\turbomed.exe") Then
   TMServCpt = "ANMELDR"
   TMStammVk = "\turbomed"
   TMStammV = "\\" & TMServCpt & TMStammVk ' "\\Linux1\Turbomed"
   Dokumente = "\\ANMELDR\Turbomed\Dokumente"
   Call fgetgsreg("ANMELDR")
  ElseIf FileExists("\\ANMELDR1\turbomed\programm\turbomed.exe") Then
   TMServCpt = "ANMELDR1"
   TMStammVk = "\turbomed"
   TMStammV = "\\" & TMServCpt & TMStammVk ' "\\Linux1\Turbomed"
   Dokumente = "\\ANMELDR1\Turbomed\Dokumente"
   Call fgetgsreg("ANMELDR1")
  ElseIf Cpt = "GSNOTEBOOK" Or Cpt = "GSN2" Then
   TMStammVk = "\Turbomed"
   TMStammV = "d:\Turbomed"
   Dokumente = TMStammV & "\Dokumente"
   TMServCpt = "localhost"
'  If 1 = 1 Or obNot And Not fileexists("\\ANMELDR\TurboMed\Programm\Turbomed.exe") Then
'   Call fStSpei(HLM, "Software\TurboMed EDV GmbH\TurboMed\Current\", "RegisterPath", "E:\TMEinzel\Programm\")
'  Else
   Call fStSpei(HLM, "Software\TurboMed EDV GmbH\TurboMed\Current\", "Path", TMStammV)
   Call fStSpei(HLM, "Software\TurboMed EDV GmbH\TurboMed\Current\", "RegisterPath", TMStammV & "\programm\")
'  Call Shell(TMStammV & "\Programm\" & "ptserv32.exe")
'  SuSh TMStammV & "\Programm\" & "ptserv32.exe", , , 0
   rufauf TMStammV & "\Programm\" & "ptserv32.exe", , , , 0, 0
'  End If
' End If
   
   Call fgetgsreg(Cpt)
  End If
'  MsgBox "Achtung Notbetrieb ohne Linux1"
'  Shell (App.Path + "\Nachricht.exe")
'  SuSh App.Path + "\Nachricht.exe", , , 0, 1
  rufauf App.Path & "\Nachricht.exe", , , , 0
 End If
' Sicherheit = "\\ANMELDR\Sicherheit"
' Sicherheit = "\\ANMELDR1\Sicherheit"
 Sicherheit = "\\ANMELDRNEU\Sicherheit"
 DSi = "\\ANMELDL\TM-DASI" ' "\\ANMELDR\TM-DASI"

' Festlegen der direkten Pfade
 EigDatDirekt = EigDat
 PatDokDirekt = PatDok
 DownDirekt = Down
 GeraldDirekt = Gerald
 KothnyDirekt = Kothny
 ReadODirekt = ReadO
 ReadOKDirekt = ReadOK
 TMServCptDirekt = TMServCpt
 DokumenteDirekt = Dokumente
 FPos = 13
 If obNot = -1 Then
  Select Case Cpt
   Case "ANMELDL", "ANMELDL1"
    EigDatDirekt = EigDatAnmL '"c:\eigene Dateien alt"
    PatDokDirekt = PatDokAnmL '"c:\P"
    DownDirekt = DownAnmL
    GeraldDirekt = GeraldAnmL
    KothnyDirekt = KothnyAnmL
    ReadODirekt = ReadOAnmL
    ReadOKDirekt = ReadOKAnmL
   Case "MITTE"
    PatDokDirekt = mitteVol & "\daten\patientendokumente"
   Case "MITTE1"
    PatDokDirekt = mitteVol & "\daten\patientendokumente"
   Case "ANMELDR", "ANMELDRNEU", "ANMELDR1"
    TMServCptDirekt = arBackup & "\TurboMed" ' h:\turbomed  ' e:\turbomed
    DokumenteDirekt = arDokumente ' h:\turbomed\dokumente   ' e:\turbomed\dokumente
  End Select
 End If
 ' fi.Stand = "2b. nach Pfadfestlegung"
 Select Case Cpt
  Case "ANMELDL", "ANMELDL1"
   DSi = alDasi ' "E:\TM-Dasi"
  Case "ANMELDR", "ANMELDRNEU"
'   DSi = arDasi ' "E:\Turbomed-Dasi"
   If obSchottdorf Then
    Dim getsend$
    getsend = ProgVerz & "LABDFUE\GETSEND.INI"
' Const getsend = "C:\Programme\LaborSchottdorf\GETSEND.INI"
    Dim schottPfad$, dzl$(), dzlz%
    schottPfad = TMStammV
    dzlz = 0
'    If dir(getsend) = "" Then
     If Not FileExists(getsend) Then
'     Call Shell(App.Path & "\nachricht.exe " & "Achtung: LaborSchottdorfdatei:" & vbCrLf & "[" & getsend & "] nicht gefunden!")
'     SuSh App.Path & "\nachricht.exe " & "Achtung: LaborSchottdorfdatei:" & vbCrLf & "[" & getsend & "] nicht gefunden!", , , 0, 1
     rufauf App.Path & "\nachricht", "Achtung: LaborSchottdorfdatei:" & vbCrLf & "[" & getsend & "] nicht gefunden!", , , 0, 1
    Else
     Open getsend For Input As #335
     Do While Not EOF(335)
      ReDim Preserve dzl(dzlz)
      Line Input #335, dzl(dzlz)
      If dzl(dzlz) Like "Pfad=*" Then
       dzl(dzlz) = "Pfad=" & TMStammV & "\labor\"
'       Dim erg$
'       erg = dir(TMStammV & "\labor", vbDirectory)
'       If erg = "" Then
       If Not DirExists(TMStammV & "\labor") Then
'        Call Shell(App.Path & "\nachricht.exe " & "Achtung: LaborSchottdorfverzeichnis:" & vbCrLf & "[" & TMStammV & "\labor\" & "] nicht gefunden!")
'        SuSh App.Path & "\nachricht.exe " & "Achtung: LaborSchottdorfverzeichnis:" & vbCrLf & "[" & TMStammV & "\labor\" & "] nicht gefunden!", , , 0, 1
        rufauf App.Path & "\nachricht", "Achtung: LaborSchottdorfverzeichnis:" & vbCrLf & "[" & TMStammV & "\labor\" & "] nicht gefunden!", , , 0, 1
       End If
      End If
      dzlz = dzlz + 1
     Loop
     Close #335
     Open getsend For Output As #335
     For dzlz = 0 To UBound(dzl)
      Print #335, dzl(dzlz)
     Next
     Close #335
    End If
    DoEvents
   End If ' obSchottdorf
 End Select
 SetProgV3 = True
schluss:
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SetProgV3/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SetProgV3

Function WNetSetz()
 On Error GoTo fehler
 WNetKorr "p:", PatDok
 FPos = 22
 WNetKorr "u:", EigDat
 FPos = 24
 WNetKorr "v:", Down
 FPos = 26
 WNetKorr "x:", TMStammV
 FPos = 28
 WNetKorr "y:", Programme
 FPos = 28
 WNetKorr "s:", Kothny
 FPos = 30
 WNetKorr "t:", Gerald
 FPos = 32
 WNetKorr "r:", ReadO
 FPos = 33
 If WV < win_vista Then
 Else
  ' führt zur net use-Ausgabe: "Nicht verfgb"
  #Const ShellaW = False
  #If ShellaW Then
  ShellaW doalsAd & acceu & AdminGes & " net use p: " & Chr$(34) & PatDok & Chr$(34) & " " & AdminGes & " /user:sturm /persistent:yes"
  ShellaW doalsAd & acceu & AdminGes & " net use u: " & Chr$(34) & EigDat & Chr$(34) & " " & AdminGes & " /user:sturm /persistent:yes"
  rufauf "cmd", "/c net use u: """ & EigDat & """ """ & AdminGes & """ /user:sturm /persistent:yes", 2, 0, -1, 0
  ShellaW doalsAd & acceu & AdminGes & " net use v: " & Chr$(34) & Down & Chr$(34) & " " & AdminGes & " /user:sturm /persistent:yes"
  ShellaW doalsAd & acceu & AdminGes & " net use x: " & Chr$(34) & TMStammV & Chr$(34) & " " & AdminGes & " /user:sturm /persistent:yes"
  ShellaW doalsAd & acceu & AdminGes & " net use y: " & Chr$(34) & Programme & Chr$(34) & " " & AdminGes & " /user:sturm /persistent:yes"
  ShellaW doalsAd & acceu & AdminGes & " net use s: " & Chr$(34) & Kothny & Chr$(34) & " " & AdminGes & " /user:sturm /persistent:yes"
  ShellaW doalsAd & acceu & AdminGes & " net use t: " & Chr$(34) & Gerald & Chr$(34) & " " & AdminGes & " /user:sturm /persistent:yes"
  ShellaW doalsAd & acceu & AdminGes & " net use r: " & Chr$(34) & ReadO & Chr$(34) & " " & AdminGes & " /user:sturm /persistent:yes"
  #End If
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WNetSetz/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'WNetSetz

Function Links()
 On Error GoTo fehler
 KWn "office.mdb.lnk", "\\linux1\gerald\", Favor
 KWn "Blutdruckdurchschnitt.xls.lnk", uVerz, Favor
 'KWn "Bayerische Euro-Gebuehrenordnung.chm", "\\linux1\daten\eigene Dateien\Abrechnung\Bayerische Euro-Gebührenordnung 1-2010", ProgVerz
 LinkErstellen "EBM", Favor, "BEuGO.chm", "\\linux1\daten\eigene Dateien\Abrechnung"
 LinkErstellen "Vergütungsübersicht", Favor, "KVB-Uebersicht-Verguetung-Diabetes.pdf", "\\linux1\daten\eigene Dateien\Abrechnung"
' LinkErstellen "Telefonnummer indentifizieren", Favor, "klickident.exe", ProgVerz & "\klickIdent Frühjahr 2005", ProgVerz & "\klickIdent Frühjahr 2005"
 KWn "Pumpeneinstellung.xls.lnk", uVerz & "Webseite Praxis 1", Favor
 KWn "Pumpeneinstellung 12-18.xls.lnk", uVerz & "Webseite Praxis 1", Favor
 KWn "Pumpeneinstellung 6-11.xls.lnk", uVerz & "Webseite Praxis 1", Favor
 KWn "Pumpeneinstellung 1-5.xls.lnk", uVerz & "Webseite Praxis 1", Favor
 KWn "Accu-Chek Smart Pix Software.lnk", pDatenb & "\Roche Diagnostics\Accu-Chek Smart Pix Software", Favor
 KWn "EDV-Anleitungen.doc.lnk", uVerz, Favor
 KWn "Merkblatt Fußsyndrom.doc.lnk", uVerz, Favor
 KWn "Onlinebefunde Labor Staber LS1Mad48.url", Gerald, Favor
 KWn "Clarity GSchade.url", Gerald, Favor
 KWn "Diabetes seit.lnk", "U:\programmierung\VS08\Projects\DiabDauer\Release", Favor
 If WV < win_vista Then KWn "Bootmenü.exe", uVerz & "Programmierung\Partitionen", StartMen
 KWn "BE-Berechnung.lnk", uVerz & "dm", Favor
 KWn "Omnipod-Gutachten.lnk", uVerz & "DM\mylife Gutachten-Assistent V4-2011-03", Favor
  
 KWn "KeePass 2 Schade" & ".lnk", "Y:", StartMenProg ' Kopiere wenn neuer
 KWn "KeePass 2 Praxis" & ".lnk", "Y:", StartMenProg ' Kopiere wenn neuer
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Links/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'Links

Function NVIni()
On Error GoTo fehler
#If True Then
 Call KWn("NVIni.exe", uVerz & "Programmierung\NVIni", autoVz)
 If FileExists(autoVz & "\NVIni.exe") Then
  For Each Fl In FSO.GetFolder(autoVz).Files
   If LCase(Fl.name) Like "netzv*.lnk" Or LCase(Fl.name) Like "nverb*.lnk" Then
    FSO.DeleteFile (Fl.Path)
   End If
  Next Fl
 End If
' oder :
' KWn "NVIni.exe.lnk", uVerz & "Programmierung\NVIni", Environ("allusersprofile") & "\Startmenü\Autostart\"
#Else
 Dim runde%, NVerbLoc$, NVOrd$, NetzVerbLoc$
 NetzVerbLoc = uVerz & "Programmierung\NetzVerbind\NVerb.exe"
 ' fi.Stand = "5. nach Netzverbloc"
 
 For runde = 1 To 8
  Select Case runde
   Case 1: NVerbLoc = "\\MITTE\u"
   Case 2: NVerbLoc = "\\ANMELDL\u"
   Case 3: NVerbLoc = "\\linux1\daten\eigene Dateien"
   Case 4: NVerbLoc = "\\linserv\u"
   Case 5: NVerbLoc = "\\linmitte\u"
   Case 6: NVerbLoc = "\\ANMELDR\u"
   Case 7: NVerbLoc = "\\ANMELDRNEU\u"
   Case 8: NVerbLoc = "\\ANMELDR1\u"
  End Select
  If FSO.FolderExists(NVerbLoc) Then
   NVOrd = NVerbLoc & "\Programmierung\Netzverbind"
   NVerbLoc = NVOrd & "NVerb.exe"
   Dim NVerbDAt As Date, NetzVerbDat As Date, Nerg%, NetzErg%
   Nerg = FileExists(NVerbLoc)
   NetzErg = FileExists(NetzVerbLoc)
   If Nerg Then NVerbDAt = FileDateTime(NVerbLoc)
   If NetzErg Then NetzVerbDat = FileDateTime(NetzVerbLoc)
   On Error Resume Next
   If NVerbDAt < NetzVerbDat Then
    Call KopDat(NetzVerbLoc, NVerbLoc)
   ElseIf NVerbDAt > NetzVerbDat Then
    Call KopDat(NVerbLoc, NetzVerbLoc)
   End If
   On Error GoTo fehler
   Call KWn("nverb.exe", uVerz & "programmierung\netzverbind", NVOrd)
   Select Case runde
    Case 1
     Call LinkErstellen("NVerbNotMitte", autoVz, "NVerb.exe", Left(NVerbLoc, Len(NVerbLoc) - Len("NVerb.exe")))
    Case 2
     Call LinkErstellen("NVerbNotANMELDL", autoVz, "NVerb.exe", Left(NVerbLoc, Len(NVerbLoc) - Len("NVerb.exe")))
    Case 3
     Call LinkErstellen("NetzVerbind", autoVz, "NVerb.exe", Left(NVerbLoc, Len(NVerbLoc) - Len("NVerb.exe")))
    Case 4
     Call LinkErstellen("NVerbNotLinserv", autoVz, "NVerb.exe", Left(NVerbLoc, Len(NVerbLoc) - Len("NVerb.exe")))
    Case 5
     Call LinkErstellen("NVerbNotLinmitte", autoVz, "NVerb.exe", Left(NVerbLoc, Len(NVerbLoc) - Len("NVerb.exe")))
    Case 6
     Call LinkErstellen("NVerbNotServer", autoVz, "NVerb.exe", Left(NVerbLoc, Len(NVerbLoc) - Len("NVerb.exe")))
   End Select
  End If
 Next runde
' Stop
#End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in NVIni/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'NVini

Sub Main()
' Dim aktfi As New FürIcon
 On Error GoTo fehler
' aktfi.Show
 If True Then
' Do While True
' cReg.ClassKey = HKEY_LOCAL_MACHINE
' cReg.SectionKey = "SOFTWARE\Classes\oeffneverz"
' cReg.CreateKey
' Schlüssel, um vom PLZ aus den Explorer mit den Dateien des Patienten öffnen zu können
 Call cReg.WriteKey("Open Folder Protocol", "", "SOFTWARE\Classes\oeffneverz", HKEY_LOCAL_MACHINE)
 Call cReg.WriteKey(" ", "URL Protocol", "SOFTWARE\Classes\oeffneverz", HKEY_LOCAL_MACHINE)
 Call cReg.WriteKey("Open Folder", "", "SOFTWARE\Classes\oeffneverz\shell\open", HKEY_LOCAL_MACHINE)
 Call cReg.WriteKey("cmd /c set url=%1 & call set url=%%url:oeffneverz:=%% & call start explorer p:\dok\%%url%%", "", "SOFTWARE\Classes\oeffneverz\shell\open\command", HKEY_LOCAL_MACHINE)
' Loop
' oEnvSystem.Environment("NVerb") = "1"
 ' fi.Show
 ' fi.Stand = "1.Beginn"
 Call StartBen("0")
 Call SetProgV '  WV = GetOSVersion schon dabei
 pDatenb = pVerz & "datenbanken"
 Call StartBen("1")
 meßzeit
#If obDebug Then
 Debug.Assert sayifisIDE
#End If
 FPos = 0
 If WV < win_vista Then RegOrt = HLM Else RegOrt = HCU
 Dim läuftschon%
 AdminPwd = holap("Administrator")
 AdminGes = "-u administrator -p " & AdminPwd
 läuftschon = fWertLesen(RegOrt, "SOFTWARE\GSProducts", "NVerb")
' test = GetReg(1, "AppEvents\Schemes\Apps\.Default\SystemExit\.Current", vns)
 Call fDWSpei(RegOrt, "SOFTWARE\GSProducts", "NVerb", 1)
 FPos = 1
 Call SetProgV2
 FPos = 3
 meßzeit
 ' fi.Stand = "2. nach cptname"
 Call StartBen("2")
 On Error Resume Next
 'Set FSO = New FileSystemObject ' = CreateObject("Scripting.FilesystemObject")
 'Set wsh = New IWshShell_Class ' = CreateObject("Wscript.Shell")
' Set WMIreg = GetObject("winmgmts:root\default:StdRegProv")
' Die Nicht-Linux (Not-)Version nur aufrufen, wenn die Linuxversion scheitern müßte
 Call StartBen("2b")
 If SetProgV3 = 0 Then GoTo schluss
 
 ' fi.Stand = "3. nach Laborschottdorf"
 Call StartBen("3")
 
 FPos = 14
 Open EigDatDirekt + "\NVerbProtok.txt" For Append As #19
 If Err.Number = 0 Then POk = -1
 Call StartBen("3a0")
 On Error GoTo fehler
 If POk Then
  Print #19, ""
  Print #19, "0a: " + CStr(Now)
 End If
 Call StartBen("3a1")
 DoEvents
 On Error Resume Next
' Call SuSh("cmd /c net use p: " & PatDok & " " & AdminGes & " /user:sturm & dir p:\eingelesen\2015 > c:\protokoll.txt", 1, , 5000000)
' Call SuSh("cmd / c net use p: || net use p: " & PatDok & " " & AdminGes & " /user:sturm")
' Dim nr As NETRESOURCE
'  ShellaW doalsad & acceu & AdminGes & " net use p: " & Chr$(34) & PatDok & Chr$(34) & " " & AdminGes & " /user:sturm /persistent:yes"
'   SuSh "cmd /c " & Chr$(34) & "net use p: " & PatDok & " " & AdminGes & " /user:sturm /persistent:yes" & Chr$(34), 1, , 0, 1
'   SuSh "net use p: /delete", 1, , 0, 1
'   SuSh "\\linux1\daten\down\pstools\psexec -u anmeldr1\administrator -p " & adminpwd & " net use p: \\linux1\daten\patientendokumente " & adminpwd & " /user:sturm /persistent:yes", 0, , 0, 1
#If demo Then
   SturmPwd = holap("Sturm")
   SuSh "cmd /c net use p: \\linux1\daten\patientendokumente " & SturmPwd & " /user:sturm & cmd", 1, , 0, 1
#End If
' Call SuSh("cmd /c md " & Chr$(34) & "p:\unsinn" & Chr$(34), 1, , 500, 0)
' Call SuSh("cmd /c ren Arztsuche.lnk Aerztsuche.lnk")
 Call StartBen("3a2")
' If Cpt = "SONO1" Then ' der stürz hier mit dem kompilierten Programm immer ab, auch ohne psexec, s.a.u.
'  ShellaW ("ipconfig /flushdns"), vbHide, , 1000
' Else
'  SuSh "ipconfig /flushdns", 3, , 0
  rufauf "ipconfig", "/flushdns", , , , 0, 1
' End If
' Call aktfi.fuehraus(doalsad, acceu & AdminGes & " ipconfig /flushdns")
 On Error GoTo fehler
 Call StartBen("3a3")
 FPos = 15
' If WV >= win_xp Then
  If IsThemeActive Then ' Windows - XP , -Windows-klassisch
    EnableTheming False
  End If
' Sicherheitseinstellungen
  Call sichZon
' End If
 ' fi.Stand = "3b. nach EnableTheming"
 FPos = 20
 DoEvents
 Call WNetSetz
 
 Call StartBen("3b")
 FPos = 16
 Call StartBen("3c")
 Call TurbomedHerricht
 
 FPos = 17
 Call Tasks
 
 Call StartBen("4")
 Call Links
 Call StartBen("5")
 Call NVIni

 ' fi.Stand = "6. nach Linkerstellen"
 Call StartBen("6")
 
 FPos = 18
 Call SystemParametersInfo(23, 0, 0, 0) ' Tastatur beschleunigen
 Call SystemParametersInfo(11, 31, 0, 0) ' Tastatur beschleunigen
 Call SetMenuUnderlines(0) ' setzt Unterstriche
 FPos = 20
 
 ' fi.Stand = "6b. nach ParameterInfo"
 Call StartBen("6b")

' If FSO Is Nothing Then Set FSO = New FileSystemObject
 FPos = 34
 If FSO.FolderExists("\\linux1\daten") Then
  Call WNetKorr("z:", "\\linux1\daten")
 End If
 
 FPos = 35
 Dim haerzte$
 haerzte = getHAPDF()
' KWn "Hausärzte", haerzte, Environ("userprofile") & "\Favoriten"
' LinkErstellen("NVerbNotMitte", GetEnvir("allusersprofile") & "\Startmenü\Programme\Autostart", "NVerb.exe", Left(NVerbLoc, Len(NVerbLoc) - Len("NVerb.exe")))
 On Error Resume Next
 Kill Favor & "Hausärzte.lnk"
 On Error GoTo fehler
 Call LinkErstellen("Hausärzte", Favor, FSO.GetFileName(haerzte), FSO.GetParentFolderName(haerzte), FSO.GetParentFolderName(haerzte))
 
 FPos = 36
 If POk Then Print #19, "1: " + CStr(Now)
 If POk Then Print #19, "2: " + CStr(Now)
 DoEvents
 ' fi.Stand = "7. nach WNetKorr"
 Call StartBen("7")
 
 FPos = 37
 
 If POk Then Print #19, "3: " + CStr(Now)
 DoEvents
 
 ' fi.Stand = "8. nach fStSpei"
 Call StartBen("8")
 
 FPos = 44
 Call RegManip1
 
 ' fi.Stand = "9. nach fStSpei"
 Call StartBen("9")
 
 FPos = 46
 If Cpt = "LABOR3" Then
  Shell "cmd /c xcopy p:\datenbanken\custobase.mdb c:\datenbanken\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy p:\datenbanken\Ekg\*.* c:\datenbanken\Ekg\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy p:\datenbanken\LuFu\*.* c:\datenbanken\LuFu\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy p:\datenbanken\Blutdruck\*.* c:\datenbanken\Blutdruck\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy p:\datenbanken\LuFu\*.* c:\datenbanken\LuFu\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy p:\datenbanken\lzekg\*.* c:\datenbanken\lzekg\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy c:\datenbanken\Ekg\*.* p:\datenbanken\Ekg\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy c:\datenbanken\LuFu\*.* p:\datenbanken\LuFu\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy c:\datenbanken\Blutdruck\*.* p:\datenbanken\Blutdruck\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy c:\datenbanken\LuFu\*.* p:\datenbanken\LuFu\ /d /s /y /h /r /c /k "
  Shell "cmd /c xcopy c:\datenbanken\lzekg\*.* p:\datenbanken\lzekg\ /d /s /y /h /r /c /k "
  pDatenb = "C:\Datenbanken\" ' 10.5.20
 End If
 ' \\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\custo med\Database
 Call fStSpei(HLM, "SOFTWARE\WOW6432Node\custo med\Database", "ActDatabasePath", pDatenb)
 If POk Then Print #19, "4: " + CStr(Now)
 DoEvents
  
 Call ProgInStart
 
 ' fi.Stand = "11. Nach Faxe aktualisieren"
 Call StartBen("11")
 
' FPos = 48
' ' fi.Stand = "10. nach ProgInsStartMenü"
 
' Call GlobalIniGet(Cpt)
' 9.6.08: Custo angleichen
 
 FPos = 51
 Dim TMPath$
 ' fi.Stand = "12. initialisiere: TMPath"
 Call StartBen("12")
 
 FPos = 52
 TMPath = getReg(2, "Software\TurboMed EDV GmbH\TurboMed\Current", "Path")
 ' fi.Stand = "12b. initialisiere: TMPath: " & TMPath
 Call StartBen("12b")
 
 FPos = 53
 sysdrv = LCase(Environ("systemdrive"))
 Call KWn("GDT.ini", Programme & "\custo", TMPath & "\Formulare\karteikarte")
 FPos = 531
 If ProgVerz <> "c:\programme" Or sysdrv <> "c:" Then
  ' fi.Stand = ' fi.Stand & vbCrLf & "ProgVerz: " & ProgVerz
  
 FPos = 5310
 Call fGDT(TMPath)
 FPos = 5315
  
 End If
 ' fi.Stand = "12c. nach TMPath"
 FPos = 538
 Call StartBen("12c")
 
 FPos = 54
 Call RegCusto
 FPos = 56
' Call VerzPrüf(sysdrv & "\gdt")
  ' fi.Stand = "13. Nach fdwSpei"
 Call StartBen("13")
  
 FPos = 62
 If POk Then Print #19, "6: " + CStr(Now)
 DoEvents
' Call setIfapPfad(Environ("SystemDrive") & "\ifapwin\hier")
 Call StartBen("15b")
 Call setIfapPfad(lokalTMExeV & "ifap")
' If fileexists(Environ("SystemDrive") & "\ifapwin\hier\wiamdb.exe") Then
 If FileExists(lokalTMExeV & "ifap\wiamdb.exe") Then
  Call AnheftNachAnw("IFAP Arzneimittel", "wiamdb.exe")
 End If
 If Cpt = "ANMELDR" Or Cpt = "ANMELDRNEU" Or Cpt = "ANMELDR1" Then
'  Call AnheftNachAnw("Biowin", "Biowin.exe")
  If obSchottdorf Then
   Call AnheftNachVerz("Labor Schottdorf abholen", ProgVerz & "Laborschottdorf", "LaborSchottdorf.exe")
  ElseIf obStaber Then
   Call AnheftNachVerz("Labor Staber abholen", ProgVerz & "LaborStaber", "LaborStaber.exe")
  Else
   Call AnheftNachVerz("Biowin", uVerz & "Programmierung\Biowin", "Biowin.exe")
   Call AnheftNachVerz("Biowin 5/05", ProgVerz & "BioWin 05 2005", "BioWin.exe")
  End If
 End If
 FPos = 66
 If IrfanPfad <> "" Then
  Call AnheftNachVerz("IrfanView", IrfanVerz, IrfanExe)
 End If
 Call AnheftNachVerz("Diabass Pro", ProgVerz & "DIABASS5.PRO", "diab5pro.exe")
 If Cpt = "ANMELDL" Or Cpt = "ANMELDL1" Or Cpt = "BUERO" Then
'  Call Anheften(AUP & "\Startmenü\Programme\Ulead PhotoImpact 6", "PhotoImpact 6.lnk")
'  Call AnheftNachVerz("Diabass Pro", ProgVerz & "DIABASS5.PRO", "diab5pro.exe")
  Call AnheftNachAnw("Glucoday", "Glucoday.exe")
  Call AnheftNachAnw("Camit Pro", "CamitPro.Exe")
  Dim oReg As New Registry
  Set oReg = Nothing
  oReg.ClassKey = HKEY_LOCAL_MACHINE
  oReg.SectionKey = "SOFTWARE\OneTouch"
  oReg.ValueKey = "DBPath"
  Debug.Print oReg.Value
  
  
  Call AnheftNachVerz("One Touch (Lifescan)", ProgVerz & "LifeScan\OneTouchDMSPro\Bin", "DMPro.exe")
 End If
 If POk Then Print #19, "7: " + CStr(Now)
 ' fi.Stand = "14. zwischen AnheftNachAnw"
 Call StartBen("14")
 
 DoEvents
' Call AnheftNachAnw("Zugriff auf Patienten in Access", "msaccess.exe", uverz & "zugriff.mdb")
' Call AnheftNachAnw("Quelle in Access", "msaccess.exe", uverz & "anamnese\Quelle.mdb")
 If WV < win_vista Then
  Call AnheftNachAnw("Winword", "Winword.exe")
  Call AnheftNachAnw("Excel", "Excel.exe")
 End If
' Call AnheftNachAnw("Arztsuche", "firefox.exe", "http://arztsuche.kvb.de/cargo/app/erweiterteSuche.htm", , , False)
' Call AnheftNachAnw("Telefonbuch", "firefox.exe", "http://www.telefonbuch.de")
 LinkErstellen "Patientenlaufzettel", Favor, "firefox.exe", , , "http://linux1/plz/"
 LinkErstellen "Auswahlliste für Patientenlaufzettel bearbeiten", Favor, "notepad++.exe", , , "\\linux1\php\php\datalist.html"
 LinkErstellen "KVB-Arztsuche", Favor, "firefox.exe", , , "http://arztsuche.kvb.de/cargo/app/erweiterteSuche.htm"
 LinkErstellen "Das Telefonbuch", Favor, "firefox.exe", , , "http://www.telefonbuch.de"
 LinkErstellen "MVV-Verbindung raussuchen", Favor, "firefox.exe", , , "http://www.mvv-muenchen.de"
 LinkErstellen "Bahnverbindung raussuchen", Favor, "firefox.exe", , , "http://www.bahn.de"
 LinkErstellen "Webmail Praxis 99202a59ab", Favor, "firefox.exe", , , "https://webmail.mnet-online.de/rc/?_task=mail"
 LinkErstellen "LibreView diabetologie@dachau-mail.de Zucker15_", Favor, "firefox.exe", , , "https://www1.libreview.com"
 LinkErstellen "Praxis Passwörter", Favor, "\\linux1\daten\shome\gerald\Praxis.kdbx"
'  If Cpt Like "SP*" Or Cpt = "ANMELDL" Or Cpt = "ANMELDL1" Then
'  Call AnheftNachAnw("BZ-HbA1c-Korrelation", "firefox.exe", "http://www.dachau-surf.de/diabetologie/HbA1cMittlererBZ.htm")
' End If
 LinkErstellen "Blutzucker-Hba1c-Korrelation", Favor, "HbA1c BZ DCCT ADAG.jpg", "\\linux1\daten\eigene Dateien\dm\"
 Call KWnK("Dienstplan.exe", "DP")
 Call AnheftNachVerz("Dienstplan Praxis", ProgVerz & "DP", "Dienstplan.exe")
 Call KWnK("DateiLese.exe", "Dateilesen")
 Call AnheftNachVerz("Patientendaten", ProgVerz & "Dateilesen", "DateiLese.exe")
 ' fi.Stand = "15. Nach Anheft"
 Call StartBen("15")
 
 If Cpt = "ANMELDR" Or Cpt = "ANMELDRNEU" Or Cpt = "ANMELDR1" Then
  Call KWnK("Laborschottdorf.exe", "Laborschottdorf")
  Call KWnK("LaborStaber.exe", "LaborStaber")
 End If
' Call AnheftNachAnw("Dienstplan alt", "Excel.exe", uverz & "Dienstplan\Dp.xls")
 FPos = 68
 Call AnheftNachAnw("Auto-Verbindung raussuchen", "AutoRout.exe")
 If POk Then Print #19, "8: " + CStr(Now)
 ' fi.Stand = "16 nach Auto-Verb"
 Call StartBen("16")
 
 DoEvents
 If InStr(UN, "erald") > 0 Or InStr(UN, "chade") > 0 Then
   Dim smverz$
   smverz = getReg(2, "Software\StarFinanz\StarMoney\8.0\app\", "Path")
'   Call AnheftNachAnw("StarMoney 5.0 Apo-Edition", "StartStarMoney.exe")
   If smverz <> "" Then Call AnheftNachVerz("StarMoney", smverz, "StartStarMoney.exe", smverz)
'   Dim mpverz$
'   mpverz = ProgVerz & "Mobipocket.com"
'   If mpverz <> "" Then Call AnheftNachVerz("Herold etc.", mpverz + "\Mobipocket Reader", "reader.exe", mpverz)
   Dim vsp$
   vsp = getReg(2, "SOFTWARE\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic", "ProductDir")
   Call AnheftNachAnw("Visual Basic 6.0", "vb6.exe", "", vsp, "vb6.exe") ' vsp+"VB6.EXE"
   Call AnheftNachAnw("Visual Basic 2005", "vbexpress.exe")
   Call AnheftNachAnw("Office", "msaccess.exe", "t:\office.mdb")
   Call KWnK("AdrAnzeig.exe", "AdressenAnsehen")
   Call AnheftNachVerz("Adressen und Kalender anzeigen", ProgVerz & "AdressenuKal", "AdrAnzeig.exe")
'   If Cpt <> "ANMELDR" Then
    Call AnheftNachVerz("SurfMusik 3.1", ProgVerz & "SurfMusik 3.1", "SurfMusik.exe")
'   End If
   Call KWnK("Verzeichnisseangleichen.exe", "Verzeichnissevergleichen")
   Call AnheftNachVerz("Verzeichnisse angleichen", uVerz & "Programmierung\Verzeichnissevergleichen", "Verzeichnisseangleichen.exe")
   If Cpt = "ANMELDL" Or Cpt = "ANMELDL1" Then
    Call AnheftNachVerz("Waverec", ProgVerz & "waverec", "waverec.exe")
   End If
 End If
 FPos = 70
 ' fi.Stand = "17 nach Diversem"
 Call StartBen("17")
 
 If POk Then Print #19, "9: " + CStr(Now)
 DoEvents
 If Cpt = "ANMELDL" Or Cpt = "ANMELDL1" Then
  Call fxset(alBoot) ' korrigiert 23.7.10, zuvor EigDatAnmL ' Korrigiert 11.6.09, zuvor EigDatDirekt
 End If
 DoEvents
 ' fi.Stand = "17b. Nach Anheft"
 Call StartBen("17b")
 
 FPos = 64
 ' geht nicht:
' If cpt = "ANMELDL" Or Cpt = "ANMELDL1"  Then
'  Call Anheften(AUP + "\Desktop", "Faxe, Liste der empfangenen.lnk")
' End If
 If Cpt = "ANMELDL" Or Cpt = "ANMELDL1" Or Cpt = "ANMELDR" Or Cpt = "ANMELRNEU" Or Cpt = "ANMELDR1" Then
  Call AnheftNachAnw("Fax senden (Classic Phonetools)", "PhonTool.exe")
 End If

 If Cpt = "ANMELDL" Or Cpt = "ANMELDL1" Then
  Call Anheften(uVerz & "Programmierung\verzeichnissevergleichen", "Linux1 sichern.lnk")
  Call Anheften(uVerz & "Programmierung\verzeichnissevergleichen", "ANMELDL sichern.lnk")
 End If
 ' fi.Stand = "18 nach Verzeichnissevergleichen"
 Call StartBen("18")
 
 FPos = 72
 If POk Then Print #19, "10: " + CStr(Now)
 DoEvents
' Call AnheftNachVerz("Telefonnummern identifizieren", ProgVerz & "klickIdent Frühjahr 2005", "klickIdent.exe", ProgVerz & "\klickIdent Frühjahr 2005")
 If Cpt = "ANMELDR" Or Cpt = "ANMELDRNEU" Or Cpt = "ANMELDR1" Then
  Call AnheftNachAnw("Thunderbird", "thunderbird.exe")
  If InStr(UN, "erald") > 0 Or InStr(UN, "chade") > 0 Then
   Call KWnK("sichkop2.exe", "SichKop")
  End If
'  Call AnheftNachVerz("Sicherheitskopien kurz", "progverz", "SichKop2.exe", , "kurz")
'  Call AnheftNachVerz("Sicherheitskopien lang", "progverz", "SichKop2.exe", , "lang")
'  Call AnheftNachVerz("Sicherheitskopien Optionen", "progverz", "SichKop2.exe", , "?")
 End If
 If POk Then Print #19, "11: " + CStr(Now)
 'Dim WSH As New IWshShell_Class
 'If IsNull(WSH) Then Set WSH = New IWshShell_Class
 FPos = 74

 ' fi.Stand = "19 nach office"
 Call StartBen("19")
 
' If Cpt = "SONO1" Then ' der stürz hier mit dem kompilierten Programm immer ab, auch ohne psexec, s.a.o.
'  Shell "explorer.exe p:\plz", vbMaximizedFocus
' Else
'  SuSh "explorer.exe p:\plz", 0, "p:\plz", 0, vbMaximizedFocus ' 3
' End If
 rufauf "firefox.exe", "http://linux1/plz/"
' rufauf "explorer", "p:\plz", 0, "p:\plz", 0, 3
 
 ' fi.Stand = "20 nach p:\plz"
 Call StartBen("20")
 
' If Cpt = "SONO1" Then ' der stürz hier mit dem kompilierten Programm immer ab, auch ohne psexec, s.a.o.
'  Shell "cmd /c xcopy v:\med-import\*.* ""%appdata%\med-import\*.*"" /s /y /h /r /c /k /d"
'  Shell "cmd /c copy ""%appdata%\med-import\med-import_" & Left(sysdrv, 1) & ".ini"" ""%appdata%\med-import\med-import.ini"" /y"
' Else
'  SuSh "xcopy v:\med-import\*.* ""%appdata%\med-import\*.*"" /s /y /h /r /c /k /d", 0, , 0
'  SuSh "cmd /c copy ""%appdata%\med-import\med-import_" & Left(sysdrv, 1) & ".ini"" ""%appdata%\med-import\med-import.ini"" /y", 0, , 0
' End If
#If False Then
 rufauf "xcopy", "v:\med-import\*.* """ & Environ("appdata") & "\med-import\*.*"" /s /y /h /r /c /k /d", , , , 0
 rufauf "cmd", "/c copy ""%appdata%\med-import\med-import_" & Left$(sysdrv, 1) & ".ini"" ""%appdata%\med-import\med-import.ini"" /y", , , , 0
#Else
 Call konfigmedimport
#End If
 FPos = 76
 Call WordOhneStartup
 FPos = 78
 If POk Then Print #19, "12: " + CStr(Now)
 FPos = 80
 If POk Then Print #19, "13: " + CStr(Now)
 On Error Resume Next
 Close #19
' Unload FI
' Notepad ersetzen
 If WV < win_vista Then Call notepadersetzen
 FPos = 77
 Call mmiAkt
 End If ' false
 Call IrfanAkt
 Call StartBen(Format(Now, "dd.mm.yy"))
schluss:
 ProgEnde
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Main/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): Resume schluss
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Main

Function konfigmedimport()
 On Error GoTo fehler
 Call VerzPrüf(Environ("appdata") & "\med-import") ', True)
 Open Environ("appdata") & "\med-import\med-import.ini" For Output As #18
 Print #18, "[setup]"
 Print #18, "lastupdate = 1444036275"
 Print #18, "EXPORT_GDT = 1"
 Print #18, "EXPORT_EMAIL = 0"
 Print #18, "EXPORT_ACM = 1"
 Print #18, "EXPORT_PDF = 1"
 Print #18, "EXPORT_CAMIT = 0"
 Print #18, "EXPORT_XLS = 1"
 Print #18, "EXPORT_CSV = 1"
 Print #18, "EXPORT_XML = 1"
 Print #18, "ShowIntro = 0"
 Print #18, "UseMeterUnit0 = 0"
 Print #18, "Exportfolder = p:\"
 Print #18, "ONTOP = 1"
 Print #18, "UseMeterPatSettings = 1"
 Print #18, "ShowPATIENT = 1"
 Print #18, "ShowCOMPLETE = 1"
 Print #18, "ShowSELECTACTION = 0"
 Print #18, "ShowSUMMARY = 1"
 Print #18, "EXPORT_000 = 0"
 Print #18, "EXPORT_GLUCOPRINT = 0"
 Print #18, "Favorites = |149|92|202|99|109|217|139|99999|100000|100001|1008|1015|113|205|251|248|203|106|239|246|175|265|266"
 Print #18, "EXPORT_SMARTPIX = 1"
 Print #18, "[GDT]"
 Print #18, "DostextImport = 0"
 Print #18, "PATH = " & Left$(sysdrv, 1) & ":\gdt\"
 Print #18, "ExtAbbrev = TURB"
 Print #18, "Importfilename = turbmedi.GDT"
 Print #18, "Exportfilename = mediturb.GDT"
 Print #18, "ExtID = Turbomed"
 Print #18, "DostextExport = 0"
 Print #18, "ExportFile = 1"
 Print #18, "ExportData = 0"
 Print #18, "ExportFileDesc = Blutzucker (@IMPORT_DATERANGE@, MBG: @IMPORT_BG_AVG@)"
 Print #18, "ExportComment = 1"
 Print #18, "[DATADEF]"
 Print #18, "DisplayUnit0 = 1"
 Print #18, "Targetfrom0 = 80"
 Print #18, "TARGETUNTIL0 = 140"
 Print #18, "Hypo0 = 60"
 Print #18, "Hyper0 = 180"
 Print #18, "[LICENSE]"
 Print #18, "1 = 5007-930001-020CD071-5800"
 Print #18, "2 = 500B-780001-0306A0CB0F8-6800"
 Print #18, "3 = 5006-870001-0110A-F438"
 Print #18, "4 = 5003-ED0001-01104-FF8F"
 Print #18, "5 = 5005-6D0001-01109-320B"
 Print #18, "6 = 5003-560001-020B90BA-C6D3"
 Print #18, "[CAMIT]"
 Print #18, "Path = pathimport"
 Print #18, "[Email]"
 Print #18, "DataFormat = 1"
 Print #18, "BODY = Diese Mail wurde mittels med-import versendet"
 Print #18, "ZIP = 1"
 Print #18, "SUBJECT = Datenversand von med-import"
 Print #18, "[USERDATA]"
 Print #18, "STREET = Mittermayerstraße 13"
 Print #18, "PHONE = 8131616380#"
 Print #18, "EMAIL = diabetologie@dachau-mail.de"
 Print #18, "City = Dachau"
 Print #18, "PRAXIS = Diabetologische Gemeinschaftspraxis Dachau"
 Print #18, "NAME = Gerald Schade"
 Print #18, "Fax = 8131616381#"
 Print #18, "ZIP = 85221"
 Print #18, "[pdf]"
 Print #18, "AttachFiles = 1"
 Print #18, "FILENAME = %name%,%firstname%,%y%%m%%d%%h%%i%%s%.PDF"
 Print #18, "PATH = p:\"
 Print #18, "Open = 0"
 Print #18, "Interval = 1"
 Print #18, "[205]"
 Print #18, "Com = 16"
 Print #18, "[CSV]"
 Print #18, "FILENAME = %name%,%firstname%%y%%m%%d%%h%%i%%s% CSV.csv"
 Print #18, "PATH = p:\"
 Print #18, "Open = 0"
 Print #18, "[xls]"
 Print #18, "FileName = %name%,%firstname% %y%%m%%d% xls.xls"
 Print #18, "PATH = P:\"
 Print #18, "Open = 0"
 Print #18, "[XML]"
 Print #18, "FILENAME = %name%,%firstname%%y%%m%%d%%h%%i%%s% XML.xml"
 Print #18, "PATH = p:\"
 Print #18, "Open = 0"
 Print #18, "[265]"
 Print #18, "Com = 301"
 Print #18, "[239]"
 Print #18, "Com = 301"
 Print #18, "[202]"
 Print #18, "Com = 4"
 Print #18, "[266]"
 Print #18, "Com = 301"
 Print #18, "[251]"
 Print #18, "Com = 13"
 Print #18, "[217]"
 Print #18, "Com = 301"
 Print #18, "[203]"
 Print #18, "Com = 301"
 Print #18, "[106]"
 Print #18, "Com = 7"
 Print #18, "[260]"
 Print #18, "Com = 301"
 Print #18, "[ACM]"
 Print #18, "FILENAME = %name%,%firstname%,%y%%m%%d%%h%%i%%s%ACM.csv"
 Print #18, "PATH = p:\"
 Print #18, "FORCEUNIT = 0"
 Close #18
  Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in konfigmedimport/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): Exit Function
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
 End Function ' konfigmedimport

Function ProgInStart&()
 On Error GoTo fehler
' Call ProgInsStartMenü("Not - Normal umschalten", "NotNormal.exe", "NotNormalbetrieb")
' Call ProgInsStartMenü("BDT-Datei aus Turbomed lesen", "Dateilese.exe", "Dateilesen")
 Call ProgInsStartMenü("Patientendaten", "Dateilese.exe", "DateiLesen")
 Call ProgInsStartMenü("Fotos benennnen", "Fotos benennen.exe", "FotosBenenn")
 Call ProgInsStartMenü("Zwei Verzeichnisse von Unterschieden bereinigen", "VZAntitwin.exe", "VerZAntitwin")
 Call ProgInsStartMenü("Code finden", "BasFinden.exe", "BasSuch")
 Call ProgInsStartMenü("DateiLesen im Debugger", "DateiLese.vbp", "DateiLesen")
 Call ProgInsStartMenü("KV-Ärzte im Debugger", "KV-Ärzte.vbp", "Hausärzte")
 Call ProgInsStartMenü("Outlook sortieren", "OutlookSortieren.exe", "OutlookSortieren")
 Call ProgInsStartMenü("BDT komprimieren", "BDTkompr.exe", "BDTkompr")
 Call ProgInsStartMenü("Duplikate löschen", "DuplikateLöschen.exe", "DuplikateLöschen")
 Call ProgInsStartMenü("Musik komprimieren", "Musik.exe", "Musik")
 Call ProgInsStartMenü("TxtZuMySQL", "TxtZuMySQL.exe", "TxtZuMySQL")
 Call ProgInsStartMenü("testAdr", "testAdr.exe", "testAdr")
 Call ProgInsStartMenü("AcKnack", "AcKnack.exe", "AcKnack")
 Call ProgInsStartMenü("TMKnack", "TMKnack.exe", "TMKnack")
 Call ProgInsStartMenü("Umlautkorrektur im Turbomedverzeichnis", "Umlautkorrektur.exe", "Umlautkorrektur")
 Call ProgInsStartMenü("Überweiserliste aktualisieren", "HAAkt.exe", "HAAkt")
 If Cpt = "ANMELDL" Or Cpt = "ANMELDL1" Or Cpt = "SONO" Then
  Call ProgInsStartMenü("DEFF-Reader (MOD-Laufwerk einlesen)", "Deff-Reader.exe", "Deff-Reader")
 End If
 If Left$(Cpt, 7) = "ANMELDL" Then
  Call ProgInsStartMenü("Faxe schnell aktualisieren", "faxakt.exe", "FaxAkt", "nurneue")
  Call ProgInsStartMenü("Faxe ausführlich aktualisieren", "faxakt.exe", "FaxAkt", "")
 End If
 ProgInStart = True
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ProgInStart/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): Exit Function
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ProgInStart

Function RegCusto&()
 On Error GoTo fehler
 Call fStSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ExportFileName", "custturb")
 Call fStSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ImportFileName", "turbcust")
 Call fStSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ExportFileNameExt", "gdt")
 Call fStSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ImportFileNameExt", "gdt")
 Call fStSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ExportFilePath", sysdrv & "\gdt")
 Call fStSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ImportFilePath", sysdrv & "\gdt")
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "AutoReturn", 1)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "Export", 1)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ExportFileExtAutoIncr", 0)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "Exportkonvertierung", 0)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ExportNachfrage", 1)
 FPos = 55
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ExportTyp", 2)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ExternalStorage", 0)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "Import", 1)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ImportFileExtAutoIncr", 0)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ImportFilePolling", 0)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "Importkonvertierung", 0)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "Importkonvertierung", 0)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "ImportTyp", 2)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "PatientenNeuaufnahme", 1)
 Call fDWSpei(HLM, "SOFTWARE\custo med\PraxisEdv", "PraxEdvTyp", 0)
 RegCusto = True
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegCusto/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): Exit Function
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' RegCusto

Function RegManip1&()
 Dim OV$
 On Error GoTo fehler
 Call fStSpei(HCU, "SOFTWARE\mediaspects\DIABASS5 PRO", "DataDir", "P:\datenbanken\DIABASS5.PRO\patients")
 Call fStSpei(HCU, "SOFTWARE\mediaspects\DIABASS5 PRO", "PatientData", "P:\datenbanken\DIABASS5.PRO\patients\")
 Call fStSpei(HCU, "AppEvents\Schemes\Apps\.Default\SystemExit\.Current", "", "Windows XP-Sprechblase.wav")
 Call fStSpei(HCU, "AppEvents\Schemes\Apps\.Default\SystemStart\.Current", "", "Windows XP-Sprechblase.wav")
 Call fStSpei(HCU, "AppEvents\Schemes\Apps\.Default\WindowsLogoff\.Current", "", "Windows XP-Sprechblase.wav")
 Call fStSpei(HCU, "AppEvents\Schemes\Apps\.Default\WindowsLogon\.Current", "", "Windows XP-Sprechblase.wav")
 Call fStSpei(HCU, "AppEvents\Schemes\Apps\.Default\SystemHand\.Current", "", "") ' beim Neuöffnen eines Explorerfenster
 Call fStSpei(HCU, "AppEvents\Schemes\Apps\Explorer\ActivatingDocument\.Current", "", "")
 Call fStSpei(HCU, "AppEvents\Schemes\Apps\Explorer\MoveMenuItem\.Current", "", "start.wav")
 FPos = 40
 Call fStSpei(HLM, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "DEVMGR_SHOW_DETAILS", "1") ' Detailansicht im Gerätemanager
 FPos = 42
 
 OV = OfficeVersion
 If OV <> ".0" Then
  Call fStSpei(HCU, "Software\Microsoft\Office\" + OV + "\Common\General", "SharedTemplates", uVerz & "Vorlagen")
  Call fStSpei(HCU, "Software\Microsoft\Office\" + OV + "\Common\General", "UserTemplates", uVerz & "Vorlagen")
 End If
 
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "ForceClassicControlPanel", 1) ' Systemsteuerung klassisch anzeigen
 Call fStSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\AutoComplete", "AutoSuggest", "no") ' sonst können Web-Adressen kaum eingegeben werden
 If POk Then Print #19, "Computer: " + CStr(Cpt) + ", Betriebssystem: " + CStr(WV) + ", Benutzer: " + CStr(UN) + ", Pfad: " + App.Path
 If POk Then Print #19, "0b: " + CStr(Now)
 
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer", "AltTabSettings", 0) ' geändert 8.9.25, statt 1
 Call fDWSpei(HCU, "Control Panel\Desktop", "CoolSwitchColumns", 10)
 Call fDWSpei(HCU, "Control Panel\Desktop", "CoolSwitchRows", 10)
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\AltTab", "Columns", 12)
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\AltTab", "Rows", 12)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer", "ThumbnailSize", 32) '
' die ersten angeblich für XP klassisch
' Call fStSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "CascadeControl Panel", "YES") ' Systemsteuerung zeigen
' Call fStSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "CascadeMyDocuments", "YES") '
' Call fStSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "CascadeMyPictures", "YES") '
 Call fStSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "CascadeNetworkConnections", "YES") '
' Call fStSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "CascadePrinters", "YES") '
' Call fStSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuScrollPrograms", "No") ' kein Bildlauf in Programme
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "IntelliMenus", 0) ' Angepasste Ausklappmenüs nicht/verwenden
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "FolderContentsInfoTip", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "FriendlyTree", 1) '
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "MapNetDrvBtn", 0) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ServerAdminUI", 0) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowCompColor", 1) '
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowInfoTip", 1) '
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_AdminToolsTemp", 2) ' Verwaltung in Alle Programme und Startmenü anzeigen
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_AutoCascade", 1) '
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_EnableDragDrop", 1) ' Ziehen und Ablegen aktivieren
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowControlPanel", 1) 'Systemsteuerung anzeigen
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowMyComputer", 1) ' Arbeitsplatz
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowMyDocs", 1) ' Eigene Dateien
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowMyMusic", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowMyPics", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowNetConn", 1) ' Netzwerkverbindungen
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowOEMLink", 1) ' Herstellereintrag
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowRecentDocs", 1) '
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartButtonBalloonTip", 2) '
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuChange", 0) ' Ziehen und Ablegen deaktivieren
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuFavorites", 2) ' Favoriten anzeigen
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuInit", 2) '
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuLogoff", 0) ' Abmelden ausblenden
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuRun", 0) ' Ausführen ausblenden
' Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarAnimations", 1) '
' Tweakui
 If 1 = 1 Then Call tweakui
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer", "EnableAutoTray", 0) 'Inaktive Symbole nicht ausblenden
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowSetProgramAccessAndDefaults", 0)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_LargeMFUIcons", 0) ' Kleine Symbole verwenden
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_MinMFU", 7)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_NotifyNewApps", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ScrollPrograms", 0)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowHelp", 1) ' Hilfe und Support anzeigen
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowNetPlaces", 1) ' Netzwerkumgebung
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowPrinters", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowRun", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowSearch", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuAdminTools", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarGlomming", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarSizeMove", 0) '1?
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "DisableThumbnailCache", 0)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ClassicViewState", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "DontPrettyPath", 0)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Filter", 0)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "FriendlyTree", 1) ' 0?
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 0)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideIcons", 0)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "NoNetCrawling", 0)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "PersistBrowsers", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SeparateProcess", 0) ' 1?
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowCompColor", 1) ' 0?
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "WebView", 1) ' 0?
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "WebViewBarricade", 1)
 Call fDWSpei(HCU, "Software\Microsoft\Office\9.0\Common\Toolbars", "AdaptiveMenus", 0)
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Start_LargeMFUIcons", 0) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Start_MinMFU", 30) '
' Call fdwspei(HCU, "SessionInformation", "ProgramCount", 11)
' Internet-Zeitserver
 Call fStSpei(HLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers", "3", "clock.isc.org") '
 Call fStSpei(HLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers", "4", "timekeeper.isi.edu") '
 Call fStSpei(HLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers", "5", "usno.pa-x.dec.com") '
 Call fStSpei(HLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers", "6", "tock.usno.navy.mil") '
 Call fStSpei(HLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers", "7", "tick.usno.navy.mil") '
 Call fStSpei(HLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers", "", "3") '
 RegManip1 = True
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegManip1/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): Exit Function
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function 'RegManip1

Function fGDT(TMPath$)
  Dim gdrd$, gdr$, GDT$, zwis$
  gdrd = "GDTroh.ini"
  gdr = TMPath & "\formulare\karteikarte\" & gdrd
  GDT = TMPath & "\formulare\karteikarte\GDT.ini"
  On Error Resume Next
  Kill gdr
  On Error GoTo fehler
 FPos = 5311
'  If LenB(dir(GDT)) <> 0 Then
  If FileExists(GDT) Then
   On Error Resume Next
    Name GDT As gdr
    If Err.Number <> 0 Then
      'ShellaW doalsad & acceu & AdminGes & " cmd /e:on /c ren " & Chr$(34) & gdt & Chr$(34) & " " & Chr$(34) & gdrd & Chr$(34), vbHide, , 10000
'      SuSh "cmd /e:on /c ren " & Chr$(34) & GDT & Chr$(34) & " " & Chr(34) & Chr$(34) & gdrd & Chr$(34), 1, , 0
       rufauf "cmd", "/e:on /c ren """ & GDT & """ """ & gdrd & """", 2, , , 0
    End If
 FPos = 5312
   On Error GoTo fehler
   Open gdr For Input As #336
   On Error Resume Next
   Open GDT For Output As #334
   If Err.Number <> 0 Then ' Datei vielleicht durch anderen Benutzer schon geöffnet
    Close #336
    Exit Function
   End If
   On Error GoTo fehler
 FPos = 5313
   Do While Not EOF(336)
    Line Input #336, zwis
    zwis = REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(zwis, "c:\programme\", ProgVerz), "C:\Program Files (x86)\\", ProgVerz), "c:", sysdrv), "C:", sysdrv), "C:\Programme\", ProgVerz), "\\\", "\"), "\\", "\")
    Print #334, zwis
    DoEvents
   Loop
 FPos = 5314
   Close #336
   Close #334
  End If
  Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GDT/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): Exit Function
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function 'GDT

Function mmiAkt()
  On Error GoTo fehler
  Call SetProgV
  Dim mmiq$, mmiz$ ' , erg$
  mmiq = "v:\mmiconfig.dat"
'  erg = dir(mmiq)
'  If erg <> vNS Then
  If FileExists(mmiq) Then
   mmiz = ProgVerz & "\MMI PHARMINDEX"
'   erg = dir(mmiz, vbDirectory)
'   If erg <> vNS Then
   If DirExists(mmiz) Then
    mmiz = mmiz & "\config.dat"
    KopDat mmiq, mmiz
   End If
  End If
  Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in mmiAkt/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): Exit Function
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' mmiakt

Function IrfanAkt() ' Konfigurationsdatei für GelbeListe
 Dim DT1$, txt0$(), ender%, aufpass%, gpos%, art$, Inhalt$, ri$, i&
 On Error GoTo fehler
 ReDim txt0(0)
 Call SetProgV
 DT1 = REPLACE$(IrfanPfad, ".exe ", ".ini") 'ProgVerz & "\IrfanView\i_view32.ini"
' If dir(DT1) <> vNS Then
 If FileExists(DT1) Then
  Open DT1 For Input As #278
  Do While Not EOF(278)
   Line Input #278, txt0(UBound(txt0))
   If Left$(txt0(UBound(txt0)), 1) = "[" Then
    If txt0(UBound(txt0)) = "[Viewing]" Then
     aufpass = True
    ElseIf aufpass Then
     aufpass = False
    End If
   End If
   If aufpass Then
    gpos = InStr(txt0(UBound(txt0)), "=")
    If gpos > 1 Then
     art = Left$(txt0(UBound(txt0)), gpos - 1)
     Inhalt = Mid$(txt0(UBound(txt0)), gpos + 1)
     Select Case art
      Case "ShowFullScreen":  ri = "3": If Inhalt <> ri Then ender = True: txt0(UBound(txt0)) = art & "=" & ri
      Case "FullBackColor":   ri = "65281": If Inhalt <> ri Then ender = True: txt0(UBound(txt0)) = art & "=" & ri
      Case "FullText":        ri = "$D$F $X|$T $S|$M": If Inhalt <> ri Then ender = True: txt0(UBound(txt0)) = art & "=" & ri
      Case "Font":            ri = "Courier": If Inhalt <> ri Then ender = True: txt0(UBound(txt0)) = art & "=" & ri
      Case "FontParam":       ri = "-13|0|0|0|400|0|0|0|0|1|2|1|49|": If Inhalt <> ri Then ender = True: txt0(UBound(txt0)) = art & "=" & ri
      Case "ViewAll":         ri = "1": If Inhalt <> ri Then ender = True: txt0(UBound(txt0)) = art & "=" & ri
      Case "ShowHiddenFiles": ri = "1": If Inhalt <> ri Then ender = True: txt0(UBound(txt0)) = art & "=" & ri
      Case "FitWindowOption": ri = "4": If Inhalt <> ri Then ender = True: txt0(UBound(txt0)) = art & "=" & ri
     End Select
    End If
   End If
   ReDim Preserve txt0(UBound(txt0) + 1)
  Loop
  Close #278
  If ender <> 0 Then
   On Error Resume Next
'   Kill ProgVerz & "\IrfanView\i_view32_alt.ini"
   Kill REPLACE$(IrfanPfad, ".exe ", "_alt.ini")
   On Error GoTo fehler
   Name ProgVerz & "\IrfanView\i_view32.ini" As ProgVerz & "\IrfanView\i_view32_alt.ini"
   Open ProgVerz & "\IrfanView\i_view32.ini" For Output As #278
   For i = 0 To UBound(txt0)
    Print #278, txt0(i)
   Next i
  Close #278
 End If
 End If
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Main/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): Exit Function
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' IrfanAkt

#If False Then
Function mmiakt_alt() ' Konfigurationsdatei für GelbeListe
 Dim DT1$, txt0$, gibts%
 On Error GoTo fehler
 Call SetProgV
 DT1 = ProgVerz & "\MMI Pharmindex\config.dat"
' If dir(DT1) <> vNS Then
 If FileExists(DT1) Then
  Open DT1 For Input As #277
  Do While Not EOF(277)
   Line Input #277, txt0
   If InStrB(txt0, "gelbeliste_13.fdb") <> 0 Then
    gibts = True
    Exit Do
   End If
  Loop
  Close #277
  If gibts = 0 Then
   On Error Resume Next
   Kill ProgVerz & "\MMI Pharmindex\config_alt.dat"
   On Error GoTo fehler
   Name ProgVerz & "\MMI Pharmindex\config.dat" As ProgVerz & "\MMI Pharmindex\config_alt.dat"
   KopDat "u:\MMIconfig.dat", ProgVerz & "\MMI Pharmindex\config.dat"
  End If
 End If
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Main/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): Exit Function
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select

End Function
#End If

Function notepadersetzen()
 Const QV$ = "\\linux1\daten\down\neu\npp\"
 Dim da1 As Date, da2 As Date, runde%
 Dim DT1$, Dt2r$, Dt2$, Dt2a$
 On Error GoTo fehler
 Call SetProgV
 DT1 = ProgVerz & "\Notepad++\notepad++.exe"
' erg = dir(DT1)
' If LenB(erg) = 0 Then
 If Not FileExists(DT1) Then
  DT1 = QV & "notepad++.exe"
'  erg = dir(DT1)
 End If
' If LenB(erg) <> 0 Then
 If FileExists(DT1) Then
  da1 = FileDateTime(DT1)
  For runde = 1 To 4
   Select Case runde
    Case 1: Dt2r = Environ("windir") & "\servicepackfiles\i386\"
    Case 2: Dt2r = Environ("windir") & "\system32\dllcache\"
    Case 3: Dt2r = Environ("windir") & "\system32\"
    Case 4: Dt2r = Environ("windir") & "\"
   End Select
   Dt2 = Dt2r & "notepad.exe"
'   erg2a = dir(Dt2)
'   If LenB(erg2a) <> 0 Then
   If FileExists(Dt2) Then
    da2 = FileDateTime(Dt2)
    If da2 < da1 Then
     Dt2a = Dt2r & "notepad alt.exe"
'     erg = dir(Dt2a)
'     If LenB(erg) = 0 Then
     If Not FileExists(Dt2a) Then
      Name Dt2 As Dt2a
     End If
     KopDat DT1, Dt2
    End If ' LenB(erg)<>0
    If runde = 3 Then
     KWnV QV & "plugins", Dt2r & "plugins"
    End If
    KWn "langs.xml", QV, Dt2r
    KWn "SciLexer.dll", QV, Dt2r
    KWn "nativeLang.xml", QV, Dt2r
   End If
  Next runde
  KWn "shortcuts.xml", QV, Environ("appdata") & "\Notepad++"
  Dim lang1&, lang2&
  If LenB(Dir$(QV & "langs.xml")) <> 0 Then lang1 = FileLen(QV & "langs.xml")
  If LenB(Dir$(Environ("windir") & "\system32\langs.xml")) <> 0 Then lang2 = FileLen(Environ("windir") & "\system32\langs.xml")
  If lang2 < lang1 Then KopDat QV & "langs.xml", Environ("windir") & "\system32\langs.xml"
 End If ' LenB(erg) <> 0 Then
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in notepadersetzen/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' notepadersetzen

Function setIfapPfad$(Pfad$)
Dim i%, Sc$, zwi$, typelib$
On Error GoTo fehler
For i = 0 To 22
 Select Case i
  Case 0: Sc = "IFAP.ApplicationObject"
  Case 1: Sc = "ifap.Card"
  Case 2: Sc = "IFAP.CATCClassification"
  Case 3: Sc = "IFAP.CATCClassifications"
  Case 4: Sc = "ifap.Cave"
  Case 5: Sc = "ifap.CaveResult"
  Case 6: Sc = "IFAP.CICDClassification"
  Case 7: Sc = "IFAP.CICDClassifications"
  Case 8: Sc = "ifap.Collection"
  Case 9: Sc = "IFAP.Composition"
  Case 10: Sc = "IFAP.Compositions"
  Case 11: Sc = "ifap.Diagnose"
  Case 12: Sc = "ifap.Diagnoses"
  Case 13: Sc = "ifap.Dosage"
  Case 14: Sc = "IFAP.Indication"
  Case 15: Sc = "IFAP.Indications"
  Case 16: Sc = "IFAP.Indications"
  Case 17: Sc = "IFAP.Manufacturer"
  Case 18: Sc = "IFAP.Medicament"
  Case 19: Sc = "IFAP.Medicaments"
  Case 20: Sc = "ifap.Patient"
  Case 21: Sc = "IFAP.Prescription"
 End Select
 zwi = getReg(HCR, Sc & "\Clsid", "")
 If zwi <> "" Then
  Call fStSpei(HCR, "CLSID\" & zwi & "\LocalServer32", "", Pfad$ & "\WIAMDB.EXE")
  If typelib = "" Then
   typelib = getReg(HCR, "CLSID\" & zwi & "\TypeLib", "")
  End If
 End If
Next i
' fi.Stand = "15b. Nach Ifap"

Call fStSpei(HLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\INDEX", "UninstallString", Environ("windir") & "\IsUn0407.exe -f" & Pfad & "\Uninst.isu")
' \HKEY_LOCAL_MACHINE\SYSEM\ControlSet003\Services\Ianmanserver\Shares, Ifapwin, CSCFlags=0
'MaxUses = 4294967295#
'Path=C:\Ifapwin
'Permissions = 0
'Remark=
'Type=0

' HKEY_CLASSES_ROOT\ifap.CaveResult\Clsid,,{031FEA4F-0948-49B2-B314-D361017025E3}
' \HKEY_CLASSES_ROOT\CLSID\{1D5999E7-D03B-4894-888A-F776F3BB91D0}\LocalServer32, (Standard), y:\Ifapwin\WIAMDB.EXE
' HKEY_CLASSES_ROOT\IFAP.CICDClassifications\Clsid,,{2C0AD305-F316-44F3-A7B6-D34875984E24}
' \HKEY_CLASSES_ROOT\CLSID\{2C0AD305-F316-44F3-A7B6-D34875984E24}\LocalServer32, (Standard), y:\Ifapwin\WIAMDB.EXE
' ...{315BD99F-DDB6-4E6D-A2C6-D64E06E6AFE5}...
' ...{34A322D7-EC6B-4465-89BD-204CDEAAABE7}...
' HKEY_CLASSES_ROOT\IFAP.ApplicationObject\Clsid,,{488182D0-2931-4146-B592-5B6F8E425B2A}
' ...{488182D0-2931-4146-B592-5B6F8E425B2A}...
' HKEY_CLASSES_ROOT\ifap.Card\Clsid,, {5C4A12F2-FC7F-4FC6-B4CA-2A1AA13119D9}
' ...{60D2041C-66D4-4D13-AFE7-CFA37F24218C}...
' ...{64927406-EB58-4AFA-B060-D16A6ECBC652}...
' ...{7A5AF3B4-0DE1-462A-8FFC-B357231A98F1}...
' HKEY_CLASSES_ROOT\ifap.Cave\Clsid,,{98A7C019-9891-4CF9-AAE2-1B60C3694524}
' HKEY_CLASSES_ROOT\IFAP.CATCClassifications\Clsid,,{9C88A23C-D550-4F4A-9093-33CCF2F10C2A}
' ...{9C88A23C-D550-4F4A-9093-33CCF2F10C2A}...
' HKEY_CLASSES_ROOT\IFAP.CATCClassification\Clsid,, {A882C23C-D5AF-400F-9BE2-D2C3F136321A}
' ...{A882C23C-D5AF-400F-9BE2-D2C3F136321A}...
' ...{A8E9ACF0-21CB-4FF0-A6A0-E31AD30DC718}...
' ...{E1A7314E-D18A-4195-B30E-62F0A9DF9403}...
' ...{EB00A72B-129F-44B2-A125-430A720B0193}...
' HKEY_CLASSES_ROOT\IFAP.CICDClassification\Clsid,,{EC96C41D-BD12-4968-9F12-6F59C28CFD27}
' ...{EC96C41D-BD12-4968-9F12-6F59C28CFD27}...


 Call fStSpei(HCR, "TypeLib\" & typelib & "\1.0\0\win32", "", Pfad & "\WIAMDB.EXE")
' Call fStSpei(Hlm, "SOFTWARE\Classes\TypeLib\" & typelib & "\1.0\0\win32", "", Pfad & "\WIAMDB.EXE") ' wohl redundant durch 1 Zeile oberhalb
' \HKEY_CLASSES_ROOT\TypeLib\{86283CF8-48CA-4C3D-8633-4C90269C1F08}\1.0\HELPDIR, (Standard), y:\ifapwin
 Call fStSpei(HCU, "Software\IFAP\INDEXPRAXIS\Trace", "LOGFILE", Pfad & "\IFAPAPP.LOG")
 Call fStSpei(HCU, "Software\Microsoft\Windows\ShellNoRoam\MUICache", Pfad & "\Wiamdb.exe", "Index3")
 Call fStSpei(HLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Wiamdb.exe", "", Pfad & "\wiamdb.exe")
' \HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Wiamdb.exe, Path,c:\Ifapwin\hier
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "SetIfapPfad/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' setIfapPfad

 Function ProgInsStartMenü(name$, Datei$, Optional Uord$, Optional arg$)
  Dim Lok$
  On Error GoTo fehler
  If Uord = "" Then Uord = name
  Lok = uVerz & "programmierung\" & Uord
  If LinkErstellen(name, Lok, Datei, , Lok, arg) Then
   Call KWn(name & ".lnk", Lok, StartMenProg & "\Eigene Programme")  ' Kopiere wenn neuer
  End If
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ProgInsStartMenü/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
 End Function ' ProgInsStartMenü
 
Function WordOhneStartup()
 Dim CLSID$, LoclSvr$, Quelle$, Ziel$, QuelF, ZielF, Fil As File
 Dim ZielDatei$
 FPos = 0
 On Error GoTo fehler
 FPos = 84
 If wsh2 Is Nothing Then Set wsh2 = New IWshShell_Class
 FPos = 86
' If FSO Is Nothing Then Set FSO = New FileSystemObject
 FPos = 88
 On Error Resume Next
 CLSID = wsh2.RegRead("HKEY_CLASSES_ROOT\Word.Application\CLSID\")
 If Err.Number <> 0 Then Exit Function ' dann winword nicht installiert
 On Error GoTo fehler
 FPos = 90
 LoclSvr = wsh2.RegRead("HKEY_CLASSES_ROOT\CLSID\" + CLSID + "\LocalServer32\")
 FPos = 100
 Quelle = Trim(REPLACE(Split(LoclSvr, " /")(0), "WINWORD.EXE", "Startup"))
 FPos = 102
 If FSO.FolderExists(Quelle) Then
  Set QuelF = FSO.GetFolder(Quelle)
 Else
'  Shell ("cmd /c copy ""%appdata%\med-import\med-import_" & Left(sysdrv, 1) & ".ini"" ""%appdata%\med-import\med-import.ini"" /y")
'   SuSh "cmd /c copy ""%appdata%\med-import\med-import_" & Left(sysdrv, 1) & ".ini"" ""%appdata%\med-import\med-import.ini"" /y", , , 0
   rufauf "cmd", "/c copy ""%appdata%\med-import\med-import_" & Left(sysdrv, 1) & ".ini"" ""%appdata%\med-import\med-import.ini"" /y", , , , 0
  Exit Function
 End If
 FPos = 104
 Ziel = REPLACE(Quelle, "Startup", "StartUpNicht")
 FPos = 106
 If FSO.FolderExists(Ziel) Then
  Set ZielF = FSO.GetFolder(Ziel)
 Else
  Set ZielF = FSO.CreateFolder(Ziel)
 End If
 FPos = 108
 For Each Fil In QuelF.Files
  FPos = 110
  ZielDatei = Ziel + "\" + Fil.name
  FPos = 112
  On Error Resume Next
  If FileExists(ZielDatei) Then
   Call FSO.DeleteFile(Fil.Path)
  Else
   Call FSO.MoveFile(Fil.Path, ZielDatei)
  End If
  On Error GoTo fehler
 Next Fil
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WordOhneStartup/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' WordOhneStartup()

Sub KWnK(D$, U$) ' Kopiere wenn neuer konstant
 On Error GoTo fehler
 Call SetProgV
' Call VerzPrüfneu(ProgVerz & U) ', True)
 Call VerzPrüf(ProgVerz & U) ', True)
 Call KWn(D, EigDatDirekt & "\Programmierung\" & U, ProgVerz & U)
 ' Bilder fürs Fax
' Call KWn("164.ico", EigDatDirekt & "\programmierung\icons\tele", ProgVerz & U)
' Call KWn("131.ico", EigDatDirekt & "\programmierung\icons\tele", ProgVerz & U)
' Call KWn("156.ico", EigDatDirekt & "\programmierung\icons\tele", ProgVerz & U)
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in KWnK/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' KWnK(D$, U$) ' Kopiere wenn neuer konstant
Sub KWnV(v1$, v2$) ' Kopiere wenn neuer Verzeichnis
 Dim erg$, Fil, Fold
 On Error GoTo fehler
 If Right(v1, 1) <> "\" Then v1 = v1 & "\"
 If Right(v2, 1) <> "\" Then v2 = v2 & "\"
 For Each Fold In FSO.GetFolder(v1).SubFolders
  KWnV v1 & Fold.name, v2 & Fold.name
 Next Fold
 For Each Fil In FSO.GetFolder(v1).Files
  KWn Fil.name, v1, v2
 Next Fil
' erg = dir(v1, vbDirectory)
' Do While LenB(erg) <> 0
'  If erg <> "." And erg <> ".." Then
'   If (GetAttr(v1 & erg) And vbDirectory) = vbDirectory Then
'    KWnV v1 & erg, v2 & erg
'   Else
'    KWn erg, v1, v2
'   End If
'  End If
'  On Error Resume Next
'  erg = Dir
'  If Err.Number <> 0 Then erg = vNS
'  On Error GoTo fehler
' Loop
Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos) & vbCrLf & "Verzeichnis 1:" & v1 & vbCrLf & "Verzeichnis 2:" & v2, vbAbortRetryIgnore, "Aufgefangener Fehler in KWnV/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' KWnV(v1$, v2$) ' Kopiere wenn neuer Verzeichnis

Sub KWn(D$, ByVal v1$, ByVal v2$) ' Kopiere wenn neuer
 'Dim FSO As New FileSystemObject
 Dim D1, D2, voll1$, voll2$, obKopier%
 On Error GoTo fehler
 If WV >= win_vista Then
     Dim Result&, AA$
     AA = Space$(255)
     Result = GetShortPathName(v1, AA, Len(AA))
     If WV < win_vista Then v1 = Mid$(AA, 1, Result)
     AA = Space$(255)
     VerzPrüf (v2)
     Result = GetShortPathName(v2, AA, Len(AA))
     If Result = 0 Then
'      Shell "runas /user:administrator ""cmd /c md """ & v2 & """ "" "
'      ShellaW doalsad & acceu & AdminGes & " cmd /e:on /c md """ & v2 & """", vbHide, , 10000
'      SuSh "cmd /e:on /c md """ & v2 & """", 1, , 0
       rufauf "cmd", "/e:on /c md """ & v2 & """", 2, , , 0
      AA = Space$(255)
      Result = GetShortPathName(v2, AA, Len(AA))
     End If
     If Result = 0 Then
      Dim werklverz$, wvlen%
      Dim Wvma$
      werklverz = userprof & "\werkl"
      Wvma = Left$(userprof, 2) & Chr$(34) & Mid(werklverz, 3) & Chr$(34) & "\zeigkurz.bat"
      wvlen = Len(werklverz)
      VerzPrüf (werklverz)
      If 0 Then
'       Shell "cmd /c md """ & werklverz & """"
'        SuSh "cmd /c md """ & werklverz & """", , , 0
      End If
'      Shell "cmd /c del """ & werklverz & "\aktpfad.txt"""
'      SuSh "cmd /c del """ & werklverz & "\aktpfad.txt""", 2, , 0
       rufauf "cmd", "/c del """ & werklverz & "\aktpfad.txt""", , , , 0
'      ShellaW "cmd /c echo @echo %~s1 ^> " & Chr$(34) & werklverz & "\aktpfad.txt" & Chr$(34) & " > " & Chr$(34) & werklverz & "\zeigkurz.bat" & Chr$(34), vbHide, , 10000
'      SuSh "cmd /c echo @echo %~s1 ^> " & Chr$(34) & werklverz & "\aktpfad.txt" & Chr$(34) & " > " & Chr$(34) & werklverz & "\zeigkurz.bat" & Chr$(34), 2, , 0
       rufauf "cmd", "/c echo @echo %~s1 ^> """ & werklverz & "\aktpfad.txt""" & " > """ & werklverz & "\zeigkurz.bat""", , , , 0
'      ShellaW ("cmd /c " & Wvma & "\zeigkurz.bat" & " """ & v2 & """"), vbHide, , 10000
'      SuSh "cmd /c " & Wvma & "\zeigkurz.bat" & " """ & v2 & """", 2, , 0
'       rufauf "cmd", "/c " & Wvma & "\zeigkurz.bat" & " """ & v2 & """", , , , 0
       rufauf "cmd", "/c " & Wvma & " """ & v2 & """", , , , 0
'      ShellaW ("cmd /e:on /c " & Wvma & "\zeigkurz.bat" & " """ & v2 & """"), vbHide, , 10000
'      SuSh "cmd /e:on /c " & Wvma & "\zeigkurz.bat" & " """ & v2 & """", 2, , 0
'       rufauf "cmd", "/e:on /c " & Wvma & "\zeigkurz.bat" & " """ & v2 & """", , , , 0
       rufauf "cmd", "/e:on /c " & Wvma & " """ & v2 & """", , , , 0
      Open werklverz & "\aktpfad.txt" For Input As #98
      Dim Text$
      While Not EOF(98)
       Line Input #98, Text
       If Result = 0 Then Result = Len(Text)
      Wend
      Close #98
      If Left$(Text, wvlen) = werklverz Then
       MsgBox "Falscher Pfad '" & v2 & "' in NVerb"
       Exit Sub
      Else
      
      End If
     Else
    If WV < win_vista Then v2 = Mid$(AA, 1, Result)
     End If
 End If
 obKopier = 0
 voll1 = v1 & IIf(Right(v1, 1) = "\", vNS, "\") & D
 voll2 = v2 & IIf(Right(v2, 1) = "\", vNS, "\") & D
 If WV < win_vista Then Call VerzPrüf(v2 & IIf(Right(v2, 1) = "\", vNS, "\"))
 If FileExists(voll1) Then
  Set D1 = FSO.GetFile(voll1)
  If FileExists(voll2) Then
   Set D2 = FSO.GetFile(voll2)
   On Error Resume Next
   If D1.DateLastModified > D2.DateLastModified Then obKopier = -1
   On Error GoTo fehler
  Else
   obKopier = -1
  End If
  If obKopier Then
   If InStrB(v1, "poetaktiv") <> 0 Then
    Call GetProcessCollection(2, "poetaktiv")
   End If
   Dim d1str$
   d1str = D1
   Call KopDat(d1str, v2 & IIf(Right$(v2, 1) = "\", "", "\"))
  End If
 End If
Exit Sub
fehler0:
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos) & vbCrLf & "Datei: " & D & vbCrLf & "Verzeichnis 1:" & v1 & vbCrLf & "Verzeichnis 2:" & v2, vbAbortRetryIgnore, "Aufgefangener Fehler in KWn/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' KWn

Sub AnheftNachVerz(Link$, Verz$, Anw$, Optional AusfInPf$, Optional arg$)
 Dim transf$
 Dim gibts%
 On Error GoTo fehler
 transf = Verz
 If Right$(Verz, 1) = "\" Then Verz = Left(Verz, Len(Verz) - 1)
 If WV < win_vista Then
  If LinkErstellen(Link, Verz, Anw, Verz, AusfInPf, arg) Then
   Call Anheften(IIf(transf = "", Verz, transf), Link + ".lnk")
  End If
 Else
  Dim Linkdatei$, Text$
  Linkdatei = AppVerz & "\" & Link & ".lnk"
'  If LenB(dir(Linkdatei)) <> 0 Then
  If FileExists(Linkdatei) Then
   Dim LName$, LPfad$, lDesc$, lWD$, lArgs$
   GetShortcutInfo Linkdatei, LName, LPfad, lDesc, lWD, lArgs
   If LCase$(LPfad) = LCase$(Verz & "\" & Anw) Then gibts = True
'   Dim ergn&
'   If dir(userprof & "\lnk_parser_cmd.exe") = "" Then
''    ShellaW "xcopy /d \\linux1\daten\down\lnk_parser_cmd.exe " & Chr$(34) & userprof & Chr$(34), vbNormalFocus, , 1000000
'    SuSh "xcopy /d \\linux1\daten\down\lnk_parser_cmd.exe " & Chr$(34) & userprof & Chr$(34), 1
'   End If
''   ergn = ShellaW(doalsad & acceu & AdminGes & " cmd /e:on /c " & userprof & "\lnk_parser_cmd.exe " & Chr$(34) & Linkdatei & Chr$(34) & " > " & Chr$(34) & Linkdatei & ".txt" & Chr$(34), vbHide, , 10000)
'   ergn = SuSh("cmd /c cd " & Chr$(34) & userprof & Chr$(34) & " & " & "lnk_parser_cmd.exe " & Chr$(34) & Linkdatei & Chr$(34) & " > " & Chr$(34) & Linkdatei & ".txt" & Chr$(34), 1, Chr$(34) & userprof & Chr$(34), , 1)
'   ergn = SuSh("cd " & Chr$(34) & userprof & Chr$(34) & " & " & "lnk_parser_cmd.exe " & Chr$(34) & Linkdatei & Chr$(34) & " > " & Chr$(34) & Linkdatei & ".txt" & Chr$(34), 1, Chr$(34) & userprof & Chr$(34), , 1)
'   ergn = SuSh(" -w " & Chr$(34) & userprof & Chr$(34) & " " & Chr$(34) & userprof & "lnk_parser_cmd.exe " & Chr$(34) & " " & Chr$(34) & Linkdatei & Chr$(34) & " >> " & Chr$(34) & Linkdatei & ".txt" & Chr$(34), 1, Chr$(34) & userprof & Chr$(34), 0)
'   Do While dir(Linkdatei & ".txt") = ""
'   Loop
'   Do While FileLen(Linkdatei & ".txt") = 0
'   Loop
'   Open Linkdatei & ".txt" For Input As #202
'   Do While Not EOF(202)
'    Line Input #202, Text
'    If Left$(Text, 19) = "Local path (ASCII):" Then
'     If InStr(LCase$(Text), LCase$(Verz & "\" & Anw)) <> 0 Then
'      gibts = True
'      Exit Do
'     End If
'    End If
''    Debug.Print text
'   Loop
'   Close #202
  End If
  If Not gibts Then
   If LinkErstellen(Link, AppVerz, Anw, Verz, AusfInPf, arg) Then
'   ShellaW doalsad & acceu & AdminGes & " cmd /e:on /c move " & Chr$(34) & userprof & "\Desktop\" & Link & ".lnk" & Chr$(34) & " " & Chr$(34) & Verz & "\" & Chr$(34), vbHide, , 10000
    Do While Not FileExists(AppVerz & "\" & Link & ".lnk")
    Loop
    Do While FileLen(AppVerz & "\" & Link & ".lnk") = 0
    Loop
   End If
'   SuSh "cmd /c move " & Chr$(34) & userprof & "\Desktop\" & Link & ".lnk" & Chr$(34) & " " & Chr$(34) & Verz & "\" & Chr$(34), 1, , 0
'   rufauf "cmd", "/c move """ & userprof & "\Desktop\" & Link & ".lnk""" & " """ & Verz & "\"""
  End If
  Call Anheften(AppVerz & "\", Link & ".lnk")
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in AnheftnachVerz/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' AnheftNachVerz

' Get information about this link.
' Return an error message if there's a problem.
' Verweise: "Microsoft Shell Controls and Automatation"
Private Function GetShortcutInfo$(ByVal full_name$, ByRef name$, ByRef Path$, ByVal descr$, ByRef working_dir$, ByRef args$)
Dim shl As Shell32.Shell
Dim shortcut_path, shortcut_name As String
Dim shortcut_folder As Shell32.Folder
Dim folder_item As Shell32.folderItem
Dim lnk As Shell32.ShellLinkObject

    On Error GoTo GetShortcutInfoError

    ' Make a Shell object.
    Set shl = New Shell32.Shell

    ' Get the shortcut's folder and name.
    shortcut_path = Left$(full_name, InStrRev(full_name, "\"))
    shortcut_name = Mid$(full_name, InStrRev(full_name, "\") + 1)
    If Not Right$(shortcut_name, 4) = ".lnk" Then _
        shortcut_name = shortcut_name & ".lnk"

    ' Get the shortcut's folder.
    Set shortcut_folder = shl.NameSpace(shortcut_path)

    ' Get the shortcut's file.
    Set folder_item = _
        shortcut_folder.Items.Item(shortcut_name)
    If folder_item Is Nothing Then
        GetShortcutInfo = "Cannot find shortcut file '" & full_name & "'"
    ElseIf Not folder_item.IsLink Then
        ' It's not a link.
        GetShortcutInfo = "Die Datei '" & full_name & "' ist keine Verknüpfung."
    Else
        ' Display the shortcut's information.
        Set lnk = folder_item.GetLink
        name = folder_item.name
        descr = lnk.Description
        Path = lnk.Path
        working_dir = lnk.WorkingDirectory
        args = lnk.Arguments
        GetShortcutInfo = ""
    End If
    Exit Function

GetShortcutInfoError:
    GetShortcutInfo = Err.Description
End Function ' GetShortcutInfo

Sub AnheftNachAnw(Link$, Anw$, Optional arg$, Optional RVerz$, Optional RAnw$, Optional neu%)
   Dim Verz$, Spl$(), i%, AusfInPf$
   On Error GoTo fehler
   AusfInPf = getReg(2, "Software\Microsoft\Windows\CurrentVersion\App Paths\" + Anw, "Path")
'   If Verz = "" Or IsNull(Verz) Then
     Verz = getReg(2, "Software\Microsoft\Windows\CurrentVersion\App Paths\" + Anw, "")
     Spl = Split(Verz, "\")
     Verz = ""
     For i = 0 To UBound(Spl) - 2
      Verz = Verz + Spl(i) + "\"
     Next
     If UBound(Spl) > 0 Then
      Verz = Verz + Spl(UBound(Spl) - 1)
     End If
     If UBound(Spl) > -1 Then
      Anw = Spl(UBound(Spl))
     End If
'   End If
   If Verz = "" And RVerz <> "" Then
    Verz = RVerz
    Anw = RAnw
   End If
   If Verz <> "" And Not IsNull(Verz) Then
    On Error Resume Next
    If neu Then Kill Verz & "\" & Link & ".lnk"
    On Error GoTo fehler
    Call AnheftNachVerz(Link, Verz, Anw, AusfInPf, arg)
   End If
   Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in AnheftNachAnw/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' AnheftNachAnw

Sub Anheften(Verz7$, Datei7$)
 Dim repl$, scheitern&
 On Error GoTo fehler
If WV < win_vista Then
 Dim Shell32 As New Shell32.Shell
 Dim Folder ' As Shell32.Folder
 Dim Verb As FolderItemVerb
 Dim fItems As FolderItems
 Dim fItem() As folderItem
 Dim folderItem As folderItem
 FPos = 1001
 Set Shell32 = CreateObject("Shell.Application") 'New Shell
 FPos = 1002
 Set Folder = Shell32.NameSpace(Verz7)
 FPos = 2
 If Not Folder Is Nothing Then
 FPos = 3
  Set fItems = Folder.Items
 FPos = 4
   If Not fItems Is Nothing Then
 FPos = 5
     Set folderItem = Folder.ParseName(Datei7)
 FPos = 6
     If Not folderItem Is Nothing Then
 FPos = 7
      For Each Verb In folderItem.Verbs
 FPos = 8
       repl = REPLACE$(Verb, "&", "")
 FPos = 9
       Debug.Print repl
 FPos = 10
       If repl = "An Startmenü anheften" Then
 FPos = 11
        Call folderItem.InvokeVerb(Verb.name)
 FPos = 12
        Exit For
       End If
      Next
     Else
      Debug.Print "Datei: " + Datei7 + " für Verz: " & Verz7 & " beim Anheften nicht gefunden!"
     End If
   Else
     Debug.Print "Verzeichnis: " + Verz7 + " nicht gefunden!"
     scheitern = True
   End If
 End If
End If
#If False Then
' On Error Resume Next
 #If nichtvonwin7aufxp Then
 ' dann auch Verweis auf %windir%\system32\shell32.dll bzw. "%windir%\syswow64\shell32.dll"
 Dim objShell As Shell32.Shell '
 Dim objFolder As Shell32.Folder
 Dim objfolderItem As Shell32.folderItem
 Dim objVerb As FolderItemVerb
 #Else
 Dim objShell
 Dim objFolder ' As Shell32.Folder
 Dim objfolderItem ' As Shell32.FolderItem
 Dim objVerb ' As FolderItemVerb
 #End If
 On Error GoTo fehler
 If WV < win_vista Then
 #If nichtvonwin7aufxp Then
    Set objShell = New Shell32.Shell 'CreateObject("Shell.Application")
 #Else
  Set objShell = CreateObject("shell.application")
 #End If
   Set objFolder = objShell.NameSpace(Verz7)
   If Not objShell Is Nothing Then
    Set objFolder = objShell.NameSpace(Verz7)
    If Not objFolder Is Nothing Then
     Set objfolderItem = objFolder.ParseName(Datei7)
     If Not objfolderItem Is Nothing Then
      For Each objVerb In objfolderItem.Verbs
       repl = REPLACE(objVerb, "&", "")
       If repl = "An Startmenü anheften" Then
        Call objfolderItem.InvokeVerb(objVerb.name)
        Exit For
       End If
      Next
     Else
      Debug.Print "Datei: " + Datei7 + " für Verz: " & Verz7 & " beim Anheften nicht gefunden!"
     End If
   Else
     Debug.Print "Verzeichnis: " + Verz7 + " nicht gefunden!"
     scheitern = True
    End If
   End If
 End If
#End If
 Dim runde&
 If WV >= win_vista Or scheitern Then
   On Error GoTo fehler
'   If LenB(dir(StartMen & "\" & Datei7)) = 0 Then
   If Not FileExists(StartMen & "\" & Datei7) Then
'    If LenB(dir(Verz7 & IIf(Right$(Verz7, 1) = "\", "", "\") & Datei7)) <> 0 Then
    If FileExists(Verz7 & IIf(Right$(Verz7, 1) = "\", "", "\") & Datei7) Then
     KopDat Verz7 & IIf(Right$(Verz7, 1) = "\", "", "\") & Datei7, StartMen & "\"
     Do
'      If LenB(Dir(StartMen & "\" & Datei7)) <> 0 Then Exit Do
      If FileExists(StartMen & "\" & Datei7) Then Exit Do
      runde = runde + 1
      If runde > 10000 Then Exit Do
     Loop
'     If LenB(Dir(StartMen & "\" & Datei7)) = 0 Then
    If Not FileExists(StartMen & "\" & Datei7) Then
'   Shell (App.Path + "\nachricht.exe " + " '" & Datei7 & "' nicht gefunden")
'   SuSh App.Path + "\nachricht.exe " + " '" & Datei7 & "' nicht gefunden", , , 0, 1
     rufauf App.Path & "\nachricht.exe", "'" & Datei7 & "' für Verzeichnis: '" & Verz7 & "' nicht gefunden"
     End If
    End If
   End If
   On Error GoTo fehler
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Anheften/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Anheften

Sub WNetKorr(buch$, Ziel$, Optional sPassword$, Optional sUsername$)
 Dim cont As String * 256, runde%, erg&
 Dim udtRES As NETRESOURCE
 On Error GoTo fehler
' If Buch = "y:" Or Buch = "y" Then Stop
 For runde = 1 To 2
  cont = Space(256)
  erg = WNetGet(buch, cont, Len(cont))
  If LCase(Trim(REPLACE(cont, Chr(0), Chr(32)))) = LCase(Ziel) Then ' And sPassword = "" Then
   Exit For
  Else
   Select Case runde
'    Case 1:
'     erg = WNetCancel(Buch, -1)
'     erg = WNetAdd(Ziel, "", Buch)
'    Case 2:
'     MsgBox "Fehler beim Mappen von " + Ziel + " in " + Buch
'   End Select
 ' NETRESOURCE-Struktur füllen
  Case 1
  'If sPassword = "" Then
  If Ziel <> "" Then erg = RemoveNetworkDrive(buch, True)
   With udtRES
    .dwType = RESOURCETYPE_DISK
    .lpLocalName = buch
    .lpRemoteName = Ziel
    .dwScope = RESOURCE_GLOBALNET
    .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    .dwUsage = RESOURCEUSAGE_CONNECTABLE
   End With
    
  ' Netzlaufwerk verbinden
   erg = WNetAddConnection2(udtRES, sPassword, sUsername, 1)
'   If erg = 0 Then
'     WNetKorr = True
'   End If
' 85 = already assigned
   End Select
   Exit For
  End If
 Next
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WNetKorr/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' WNetKorr

' Laufwerk auf Existenz prüfen
Public Function DriveExists%(sDrive$)
  Dim sDrives$
  On Error GoTo fehler
    
  ' Laufwerksliste ermitteln
  sDrives = Space$(255)
  If GetLogicalDriveStrings(Len(sDrives), sDrives) Then
    ' ist der Laufwerksbuchstabe enthalten?
    DriveExists = InStr(1, sDrives, sDrive, vbTextCompare)
  End If
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DriveExists/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DriveExists

' Netzlaufwerk trennen
Public Function RemoveNetworkDrive%(sDriveLetter$, bForce%) ' sDriveLetter mit zwei Buchstaben!

  Dim nResult&
  On Error GoTo fehler
  nResult = WNetCancelConnection2(sDriveLetter, CONNECT_UPDATE_PROFILE, bForce)
  RemoveNetworkDrive = (nResult = 0)
  If nResult <> 0 And nResult <> 2250 Then ' 2250 = Netzwerkverbindung ist nicht vorhanden
'   Shell (App.Path + "\nachricht.exe " + "Fehler " + APIErrorDescription(Err.LastDllError) + " beim Lösen der Netzwerkverbindung " + sDriveLetter)
'   SuSh App.Path + "\nachricht.exe " + "Fehler " + APIErrorDescription(Err.LastDllError) + " beim Lösen der Netzwerkverbindung " + sDriveLetter, , , 0, 1
   rufauf App.Path & "\nachricht.exe", "Fehler " + APIErrorDescription(Err.LastDllError) + " beim Lösen der Netzwerkverbindung " + sDriveLetter, , , 0, 1
  End If
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RemoveNetworkDrive/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' RemoveNetworkDrive

Public Function AddNetworkDrive%(sDriveLetter$, sNetWorkPath$, Optional sUsername$, Optional sPassword$)
  Dim nResult&
  Dim udtRES As NETRESOURCE
  On Error GoTo fehler
  ' Plausi auf die Parameter
  AddNetworkDrive = False
  If Len(sDriveLetter) <> 2 Or Right$(sDriveLetter, 1) <> ":" Then Exit Function
  If Len(sNetWorkPath) < 4 Then Exit Function
    
  ' ist der Laufwerksbuchstabe schon vergeben?
  If DriveExists(sDriveLetter) Then Exit Function
    
  ' NETRESOURCE-Struktur füllen
  With udtRES
    .dwType = RESOURCETYPE_DISK
    .lpLocalName = sDriveLetter
    .lpRemoteName = sNetWorkPath
  End With
    
  ' Netzlaufwerk verbinden
  nResult = WNetAddConnection2(udtRES, sPassword, sUsername, 1)
  If nResult = 0 Then
    AddNetworkDrive = True
  End If
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in AddNetworkDrive/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AddNetworkDrive

Function MAKEINTATOM$(nm&)
  Dim s$
  Const MAX_ATOM& = 30
  s = Space$(MAX_ATOM)
  MAKEINTATOM = GetAtomName(nm, s, Len(s))
End Function ' MAKEINTATOM$

Function getStartknopf&() ' s. Unterfenster
 Dim hwnd&, RetHwnd&
 Dim ClassName As String * 256
 hwnd = GetDesktopWindow()
 RetHwnd = GetWindow(hwnd, GW_CHILD)
 RetHwnd = GetWindow(RetHwnd, GW_HWNDFIRST)
 If RetHwnd <> 0 Then
  Do
    Call GetClassName(RetHwnd, ClassName, Len(ClassName))
    If Left$(ClassName, InStr(1, ClassName, vbNullChar) - 1) = "Button" Then
     getStartknopf = RetHwnd
     Exit Function
    End If
    RetHwnd = GetWindow(RetHwnd, GW_HWNDNEXT)
  Loop Until RetHwnd = 0
 End If
End Function ' getStartknopf

Function zeichneneu()
'   erg0 = RedrawWindow(ohWnd, ClientRect, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN Or RDW_FRAME)
 Dim hwnd&, RetHwnd&
 Dim ClassName As String * 256
 hwnd = GetDesktopWindow()
 RetHwnd = GetWindow(hwnd, GW_CHILD)
 RetHwnd = GetWindow(RetHwnd, GW_HWNDFIRST)
 If RetHwnd <> 0 Then
  Do
    Call GetClassName(RetHwnd, ClassName, Len(ClassName))
    Dim ClientRect As cRECT
    GetClientRect RetHwnd, ClientRect
    RedrawWindow RetHwnd, ClientRect, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN Or RDW_FRAME
    UpdateWindow RetHwnd
    RetHwnd = GetWindow(RetHwnd, GW_HWNDNEXT)
  Loop Until RetHwnd = 0
 End If
End Function ' zeichneneu

#If False Then
' Ermittelt alle Kindfenster der Form in Z-Order-Reihenfolge von hinten nach vorn
Sub Unterfenster(hwnd&)
  Dim RetHwnd As Long, FensterKlasse As WNDCLASSEX
  Dim ClassName As String, hInstance As Long
 
  FensterKlasse.cbSize = Len(FensterKlasse)
 
  ' erstes Kindfenster in der Kindfenster-Z-Order ermitteln
  RetHwnd = GetWindow(hwnd, GW_CHILD)
  RetHwnd = GetWindow(RetHwnd, GW_HWNDFIRST)
 
  If RetHwnd <> 0 Then
    With FensterKlasse
      Do
 
        ' Klassennamen ermitteln
        ClassName = Space(256)
        Call GetClassName(RetHwnd, ClassName, Len(ClassName))
        ClassName = Left$(ClassName, InStr(1, ClassName, _
        vbNullChar) - 1)
        ' Instanz ermitteln
        hInstance = GetWindowLong(RetHwnd, GWL_HINSTANCE)
 
        ' Fensterklasseninformationen ermitteln
        Call GetClassInfoEx(hInstance, ClassName, FensterKlasse)
 
        ' Informationen über die Fensterklasse des Fensters ausgeben
        Debug.Print "Klassenname: " & ClassName
        Debug.Print "Icon Handle: " & .hIcon
        Debug.Print "Kleiner Icon Handle: " & .hIconSm
        Debug.Print "Cursor Handle: " & .hCursor
        ' Nächstes Kindfenster ermitteln
        RetHwnd = GetWindow(RetHwnd, GW_HWNDNEXT)
 
      Loop Until RetHwnd = 0
    End With
  End If
End Sub ' Unterfenster
#End If

Function StartBen(NText$)
 Const WM_GETTEXT = &HD
 Const WM_SETTEXT As Long = &HC

 Dim ohWnd As Long, erg0&
 Dim lLength As Long
 Dim sWindowText As String * 255
 On Error GoTo fehler:
 If WV < win_vista Then
  ohWnd = FindWindow("shell_traywnd", "")
  ohWnd = GetWindow(ohWnd, GW_CHILD)
  lLength = SendMessage(ohWnd, WM_GETTEXT, Len(sWindowText) + 1, ByVal sWindowText)
  SendMessage ohWnd, WM_SETTEXT, 1, ByVal NText
 Else
'  ohWnd = FindWindow("Shell_TrayWnd", "")
'  lText = "2b"
'  ohWnd = FindWindowEx(GetDesktopWindow(), 0, "Button", lText)
'  Unterfenster (GetDesktopWindow())
  ohWnd = getStartknopf
'  ohWnd = FindWindowEx(ohWnd, 0, "Button", "")
'  erg0 = SetWindowText(ohWnd, NText) ' würde auch gehen
  If ohWnd <> 0 Then
'   Const WM_SETREDRAW = &HB
'   SendMessage ohWnd, WM_SETREDRAW, -1, 0
   SendMessage ohWnd, WM_SETTEXT, 1, ByVal NText
#If 0 Then
   ' das Folgende hilft alles nix
   Call zeichneneu
   Call InvalidateRect(0&, 0&, False)
'   Const WM_PAINT = &HF
'   SendMessage ohWnd, WM_PAINT, 0, 0
'   SendMessage GetDesktopWindow(), WM_PAINT, 0, 0
   UpdateWindow GetDesktopWindow()
   UpdateWindow ohWnd
   Dim ClientRect As cRECT
   GetClientRect ohWnd, ClientRect
   erg0 = RedrawWindow(ohWnd, ClientRect, 0&, RDW_UPDATENOW)
   erg0 = RedrawWindow(ohWnd, ClientRect, 0&, RDW_ERASE Or RDW_UPDATENOW)
   erg0 = RedrawWindow(ohWnd, ClientRect, 0&, RDW_INVALIDATE)
   erg0 = RedrawWindow(ohWnd, ClientRect, 0&, RDW_UPDATENOW)
   erg0 = RedrawWindow(ohWnd, ClientRect, 0&, RDW_VALIDATE Or RDW_UPDATENOW)
   erg0 = RedrawWindow(ohWnd, ClientRect, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN Or RDW_FRAME)
   UpdateWindow GetDesktopWindow()
   GetClientRect GetDesktopWindow(), ClientRect
   erg0 = RedrawWindow(GetDesktopWindow(), ClientRect, 0&, RDW_UPDATENOW Or RDW_ALLCHILDREN Or RDW_INVALIDATE)
   erg0 = RedrawWindow(ohWnd, ClientRect, 0&, RDW_UPDATENOW)
   erg0 = RedrawWindow(ohWnd, ClientRect, 0&, RDW_VALIDATE Or RDW_UPDATENOW)
   DoEvents
#End If
  End If
 End If
'SendMessage ohWnd, WM_SETTEXT, 1, ByVal "Start"
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WNetKorr/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' StartBen

Function LinkErstellen%(LName$, LPfad$, ZName$, Optional ZPfad$, Optional AusfInPf$, Optional arg$)
' Es muss ein Verweis auf 'Microsoft Scripting Runtime' gesetzt werden.
' Es muss ein Verweis auf 'Windows Script Host Object Model' gesetzt
' werden.
'Set WSH = New IWshShell_Class ' wshom_ocx
Dim ZVoll$
Dim SCut, DestFolder As Folder
On Error Resume Next
If wsh2 Is Nothing Then Set wsh2 = New IWshShell_Class
If Err.Number <> 0 Then Exit Function
On Error GoTo fehler
If IsNull(ZPfad) Or LenB(ZPfad) = 0 Then ZPfad = getReg(2, "Software\Microsoft\Windows\CurrentVersion\App Paths\" + ZName, "Path")
If IsNull(ZPfad) Or LenB(ZPfad) = 0 Then
 ZPfad = LPfad
End If
If IsNull(AusfInPf) Or LenB(AusfInPf) = 0 Then AusfInPf = getReg(2, "Software\Microsoft\Windows\CurrentVersion\App Paths\" + ZName, "Path")
If IsNull(AusfInPf) Or LenB(AusfInPf) = 0 Then AusfInPf = ZPfad
If Right$(AusfInPf, 1) <> "\" Then AusfInPf = AusfInPf & "\"
If Mid$(ZName, 2, 1) = ":" Or Mid$(ZName, 1, 2) = "\\" Then ZVoll = ZName Else ZVoll = getReg(2, "Software\Microsoft\Windows\CurrentVersion\App Paths\" + ZName, "")
If IsNull(ZVoll) Or ZVoll = "" Then
 If Not IsNull(ZPfad) And LenB(ZPfad) <> 0 Then ZVoll = ZPfad & IIf(Right$(ZPfad, 1) = "\", "", "\") & ZName
End If
     'Falls Link nicht schon existiert
'If Not fileexists(WSH.SpecialFolders(0) & "\Test.lnk") Then
' If Not fileexists(LPfad & "\" & LName & ".lnk") And fileexists(ZPfad + "\" + ZName) Then
 Dim obFE%, fileL As File, fileZ As File
 On Error Resume Next
 Set fileL = FSO.GetFile(LPfad & "\" & LName & ".lnk")
 Set fileZ = FSO.GetFile(ZVoll)
 If fileZ Is Nothing Then ' nachricht
  Debug.Print "ZVoll: " & ZVoll & " nicht gefunden!"
'  Shell (App.Path + "\nachricht.exe " + zvoll & " nicht gefunden!")
'  SuSh App.Path + "\nachricht.exe " + zvoll & & " nicht gefunden!", , , 0, 1
   rufauf App.Path & "\nachricht.exe", "ZVoll: " & ZVoll & " nicht gefunden!", , , 0, 1
 Else
  On Error GoTo fehler
  If Not fileL Is Nothing And Not fileZ Is Nothing Then
' Verknüpfung prüfen
   Dim cS As New cShellLink
   cS.FileName = fileL
   cS.LoadLink
   If UCase$(cS.Path) <> UCase$(fileZ) Or cS.WorkingDirectory <> AusfInPf Or cS.Arguments <> arg Then
    obFE = True
   Else
    LinkErstellen = True
   End If
  End If
  If (fileL Is Nothing And Not fileZ Is Nothing) Or obFE Then
    ' WSH.SpecialFolders(0) = All User Desktop Verzeichniss
'    Set DestFolder = FSO.GetFolder(LPfad)
    ' Erstelle einen Link mit dem Namen Test.lnk
'    Set SCut = WSH.CreateShortcut(DestFolder.Path & "\" & LName)
    Set SCut = wsh2.CreateShortcut(LPfad & IIf(Right(LPfad, 1) = "\", "", "\") & LName & ".lnk")
'    If SCut.TargetPath <> ZVoll Then
     SCut.TargetPath = ZVoll
     If LenB(AusfInPf) <> 0 And Not IsNull(AusfInPf) Then
      SCut.WorkingDirectory = AusfInPf
     End If
     If LenB(arg) <> 0 And Not IsNull(arg) Then
      SCut.Arguments = arg
     End If
'     SCut.IconLocation = ZVoll
     On Error Resume Next
     SCut.Save
     If Not SCut Is Nothing Then
      LinkErstellen = True
     End If
'    End If
    Set SCut = Nothing
  End If
    ' Verknüpfe es mit der ausführbaren Datei:
    'Speichere dieses Element
    ' Entladen des Objektes

'Set FSO = Nothing
'Set WSH = Nothing
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in LinkErstellen/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LinkErstellen

Function getTMExeV$(Optional idt As TMIniDatei)
 Dim hgFarbe$, kkFarbe$, kfFarbe$, pfFarbe$
 Call SetProgV
 Dim TMServer% ' bis 7.1.24 obvirt
 If FSO.FolderExists("\\linux1\turbomed\PraxisDB") Then
  TMServer = 0 ' Linux
 ElseIf FSO.FolderExists("\\linux1\turbomed\PraxisDB-wser") Then
  TMServer = 1 ' Windows auf res1: wser
 Else ' PraxisDB-res => virtwin
  TMServer = 2
 End If
 ' blau, grün, gelb
 Select Case obNot
  Case -2: hgFarbe = "RGB ( 255 / 230 / 230 )": kkFarbe = hgFarbe ' blass
  Case Else:
   If TMServer = 0 Then
    hgFarbe = "RGB ( 184 / 254 / 197 )" '"RGB ( 0 / 183 / 183 )": kkFarbe = hgFarbe ' "RGB ( 64 / 128 / 128 )"
   ElseIf TMServer = 1 Then
    hgFarbe = "RGB ( 175 / 205 / 255 )" ' blau (je höher die ersten beiden Zahlen, desto heller)
   Else ' virtwin
    hgFarbe = "RGB ( 254 / 215 / 122 )" ' gelb - orange
   End If
 End Select
 kkFarbe = "RGB ( 255 / 255 / 225 )" ' Karteikartenfarbe
 kfFarbe = "RGB( 255 / 250 / 250 )"  ' Kassenfallfarbe
 pfFarbe = hgFarbe                   ' Privatfallfarbe
 On Error GoTo fehler
 Set idt = New TMIniDatei
' 10.4.20: gibts offenbar nicht mehr
' Call idt.SetProp("Sonstiges", "TurboMed.net: Intervall Einwahl-Verbindungsprüfung in Minuten (0:deaktiviert)", "0")
' Call idt.SetProp("Sonstiges", "TurboMed.net: Intervall Router-Verbindungsprüfung in Minuten (0:deaktiviert)", "0")
 Call idt.SetProp("Sonstiges", "Neues Hauptmenü aktivieren", "nein")
 If obSchottdorf Or obStaber Then
  Call idt.SetProp("Laborimport", "Standard-Pfad zu Laborimportdateien", "\\linux1\turbomed\Labor\labor.dat") ' TMStammV + "\Labor\labor.dat")
  Call idt.SetProp("Laborimport", "Sollen in allen Laborunterverzeichnissen Wochentagsordner angelegt werden?", "ja")
 Else
  Call idt.SetProp("Laborimport", "Standard-Pfad zu Laborimportdateien", TMNotV + "Labor\Import")
 End If
 Call idt.SetProp("Anzeige/Behandlungsfall-Übersicht", "Genuiner Arzt", "ja")
 Call idt.SetProp("Anzeige/Behandlungsfall-Übersicht", "Hausärzte", "ja")
 If False And (Cpt = "MITTE1" Or Cpt = "SZ2N1") Then ' hier kommt immer ein sql-Fehler
  Call idt.SetProp("Anzeige/Behandlungsfall-Übersicht", "Schnellübersicht 1 aktiv", "nein")
  Call idt.SetProp("Anzeige/Behandlungsfall-Übersicht", "Schnellübersicht 2 aktiv", "nein")
 Else
  Call idt.SetProp("Anzeige/Behandlungsfall-Übersicht", "Schnellübersicht 1 aktiv", "ja")
  Call idt.SetProp("Anzeige/Behandlungsfall-Übersicht", "Schnellübersicht 2 aktiv", "ja")
 End If
 Call idt.SetProp("Anzeige/Patienten", "Behandlungsfall-Auswahl", "nein")
 Call idt.SetProp("Anzeige/Patienten", "Farbliche Hervorhebung aktuell abrechnungsrelevanter Patienten", "ja")
 Call idt.SetProp("Anzeige/Patienten", "Kostenträger anzeigen", "ja")
 Call idt.SetProp("Auswahlen", "Abfrage beim Löschen", "ja")
 Call idt.SetProp("Auswahlen", "Überschriften anzeigen", "ja")
 Call idt.SetProp("Automation", "Neuaufnahme Voreinstellung Abrechnungsgebiet", "7")
 Call idt.SetProp("Automation", "Neuaufnahme Voreinstellung Behandlungsfall-Typ", "1")
 Call idt.SetProp("Automation", "Neuaufnahme Voreinstellung Scheinuntergruppe", "24")
 Call idt.SetProp("Desktopobjekte", "Inhalt von 'Notiz' beim Patienten anzeigen", "ja")
 Call idt.SetProp("Farbeinstellungen", "Hauptfenster", hgFarbe)
 Call idt.SetProp("Farbeinstellungen/Farbgebung des Patientenfensters", "Fensterfarbe", hgFarbe)
 Call idt.SetProp("Farbeinstellungen/Farbgebung der Karteikarte", "Hintergrundfarbe", kkFarbe)
 Call idt.SetProp("Farbeinstellungen/Farbgebung der Karteikarte", "Hintergrundfarbe Kassenfall", kfFarbe)
 Call idt.SetProp("Farbeinstellungen/Farbgebung der Karteikarte", "Hintergrundfarbe Privatfall", pfFarbe)
' Call iDt.SetProp("Farbeinstellungen/Farbgebung der Auswahlfenster", "Fensterfarbe", "RGB ( 120 / 170 / 170 )")
 Call idt.SetProp("Laborimport", "Eintragsdatum:", "2")
 Call idt.SetProp("Laborimport", "Labor sendet bei Endbefund: den Gesamtbefund (ja) nur Laborwertänderungen (nein)", "ja")
 Call idt.SetProp("Laborimport", "Laborimportuhrzeit (volle Stunden)", "12")
 Call idt.SetProp("Laborimport", "Möchten Sie, dass auch (Facharzt-)Laborimporte ohne Anforderungsnummern im System vorhandene Laboraufträge der betroffenen Behandlungsfälle hinsichtlich Datum und Befundart des letzten Importes aktualisieren?", "nein")
 Call idt.SetProp("Laborimport", "Sollen LDT-Importe in abgeschlossene Fälle ermöglicht werden?", "ja")
 Call idt.SetProp("Laborimport", "Standard-Pfad zu Laborimportdateien", TMStammV & "\Labor\labor.dat")
 Call idt.SetProp("Karteikarte", "EBM 2000plus Arztgruppenüberprüfung", "ja")
 Call idt.SetProp("Karteikarte", "Anzahl dargestellter Labortests pro Karteieintrag", "133")
 If WV >= win_vista Then
  IrfanVerz = ProgVerzO & "\irfanview\"
  IrfanExe = "i_view64.exe "
  IrfanPfad = IrfanVerz & IrfanExe
  IrfanErg = FileExists(IrfanPfad) ' LenB(Dir(IrfanDatei)) <> 0
  If Not IrfanErg Then IrfanPfad = ""
 End If
 If IrfanPfad = "" Then
  IrfanVerz = ProgVerz & "\irfanview\"
  IrfanExe = "i_view32.exe "
  IrfanPfad = IrfanVerz & IrfanExe
  IrfanErg = FileExists(IrfanPfad) ' LenB(Dir(IrfanDatei)) <> 0
  If Not IrfanErg Then IrfanPfad = ""
 End If
' Einschub wegen Turbomedfehler 1. Quartal 2012
 If Date < #4/1/2012# Then IrfanErg = 0
 Call idt.SetProp("TurboMed Grundeinstellungen/KVK-Leser", "KVK-Lesegerät", "6") ' MKT+
 Call idt.SetProp("TurboMed Grundeinstellungen/KVK-Leser", "Mobiler Leser im stationären Modus", "ja")
 Call idt.SetProp("TurboMed Grundeinstellungen/Programmpfade", "Bildverarbeitung", IIf(IrfanErg, IrfanPfad, ""))
 Call idt.SetProp("TurboMed Grundeinstellungen/Programmpfade", "Interne Bildverarbeitung benutzen", IIf(IrfanErg, "nein", "ja"))
 Call idt.SetProp("Verzeichnisse/TurboMed", "Lizenz zuerst auf Server suchen", "nein")
 Call idt.SetProp("Verzeichnisse/TurboMed", "Mehrplatzbetrieb", "ja")
 Call idt.SetProp("Verzeichnisse/TurboMed", "Serverbetriebssystem ist Linux", IIf(InStr(LCase$(TMServCpt), "lin") > 0 And TMServer = 0, "ja", "nein")) ' oder obnot = -1 oder -2
 Dim Srs() As Variant
 Srs = Array("linux1", "192.168.178.21", "192.168.178.46")
 Dim aktSr$
 Dim i%
 Dim clsr As New clsresolve
 For i = 0 To UBound(Srs)
  aktSr = Srs(i)
'  If obVirt Then aktSr = "virtwin" ' "192.168.178.251"
  Select Case TMServer
   Case 0: aktSr = "linux1"
   Case 1: aktSr = "wser"
   Case Else: aktSr = "virtwin" ' "192.168.178.251"
  End Select
  If doPing(clsr.GetIPFromHostName(aktSr)) = 0 Then
   Call idt.SetProp("Verzeichnisse/TurboMed", "Server", aktSr) ' TMServCpt
   Call idt.SetProp("Verzeichnisse/TurboMed", "Serverpfad", "\\" & aktSr & "\turbomed")  ' TMStammV) '"turbomed")
   Exit For
  End If ' doPing
 Next i
 
 Call idt.SetProp("TurboMed Grundeinstellungen/eGK", "eMP automatisch auf eGK schreiben", "0") ' absurde Rückfrage, 8.5.22
 Call idt.SetProp("TurboMed Grundeinstellungen/Anzeige/Kartei/Abrechnung", "In Titelzeile anzeigen - Pat.-Nr.", "ja") ' 9.5.22
 Call idt.SetProp("TurboMed Grundeinstellungen/Anzeige/Kartei/Abrechnung", "In Titelzeile anzeigen - Schwangerschaft", "ja") ' 9.5.22
 Call idt.SetProp("TurboMed Grundeinstellungen/Anzeige/Kartei/Abrechnung", "In Titelzeile anzeigen - Handynummer", "ja") ' 9.5.22
 Call idt.SetProp("TurboMed Grundeinstellungen/Anzeige/Kartei/Abrechnung", "In Titelzeile anzeigen - Telefonnummer", "ja") ' 9.5.22
 Call idt.SetProp("Verzeichnisse/TurboMed", "Pfad", idt.LokalTurbomed)
 Call idt.SetProp("Verzeichnisse/TurboMed/Vorlagen", "Pfad", TMStammV & "\" & "vorlagen") ' TMStammV + "\vorlagen") ' x:\
 Call idt.SetProp("Verzeichnisse/TurboMed/Dokumente", "Pfad", IIf(obNot = -1, TMNotV + "Dokumente", IIf(i = 3, "\\" + aktSr + "\daten\turbomed\", "") + Dokumente)) ' Ergänzung 5.8.18
 Call idt.SetProp("Verzeichnisse/TurboMed/StammDB", "Pfad", "StammDB") ', TMStammV + "\StammDB") ' 29.7.15 ' x:\
 Call idt.SetProp("Verzeichnisse/TurboMed/PraxisDB", "Pfad", "PraxisDB") ' TMStammV + "\PraxisDB")' x:\
 Call idt.SetProp("Verzeichnisse/TurboMed/DruckDB", "Pfad", "DruckDB") ' , TMStammV + "\DruckDB")' x:\
 Call idt.SetProp("Verzeichnisse/TurboMed/Dictionary", "Pfad", "Dictionary") ' , TMStammV + "\Dictionary")' x:\
 Call idt.SetProp("Verzeichnisse/TurboMed/Netzwerk-Setup", "Pfad", TMStammV + "\netsetup")
 Call idt.SetProp("Datensicherung", "Automatische Datensicherung bei Programmende?", "nein") 'IIf(Cpt = "ANMELDL", "ja", "nein"))
 Call idt.SetProp("Datensicherung", "Automatische Datenspiegelung bei Programmende?", "nein")
 Call idt.SetProp("Datensicherung", "Automatische Prüfung bei Datensicherung ohne Nachfrage", "ja")
 Call idt.SetProp("Datensicherung", "Automatische Sicherung der Dokumente bei Programmende?", "nein")
 Call idt.SetProp("Datensicherung", "Automatische Spiegelung der Dokumente bei Programmende?", "nein")
 Call idt.SetProp("Datensicherung", "Automatische Sicherung der Videodaten bei Programmende?", "nein")
 Call idt.SetProp("Datensicherung", "Löschen alter Sicherungen ohne Nachfrage", "nein")
 Call idt.SetProp("Datensicherung", "Automatische Spiegelung der Audiodaten bei Programmende?", "nein")
 Call idt.SetProp("Datensicherung", "Automatische Spiegelung der Briefe bei Programmende?", "nein")
 Call idt.SetProp("Datensicherung", "Automatische Spiegelung der Dokumente bei Programmende?", "nein")
 Call idt.SetProp("Datensicherung", "Automatische Spiegelung der externen Datenbanken bei Programmende?", "nein")
 Call idt.SetProp("Datensicherung", "Automatische Spiegelung der Videodaten bei Programmende?", "nein")
' Call idt.SetProp("Datensicherung", "Beim Löschen in den Papierkorb verschieben?", "nein")
 Call idt.SetProp("Datensicherung", "Löschen alter Sicherungen ohne Nachfrage", "nein")
 Call idt.SetProp("Datensicherung/Befunde", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Beliebige Dateien", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Datenstaemme", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Dokumente/Audio", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Dokumente/Briefe", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Dokumente", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Dokumente/Video", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Druckaufträge", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Externe Datenbanken", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Formulare", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Praxisdaten", "Zielverzeichnis für Sicherung", DSi)
 Call idt.SetProp("Datensicherung/Vorlagen", "Zielverzeichnis für Sicherung", DSi)
 Call idt.Sichern
 getTMExeV = idt.TMVerz
 Call LizKop(idt.LokalTurbomed)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in getTMExeV/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getTMExeV

Function LizKop(LokTM$)
' Überprüfen der Turbomed-Lizenzdatei
 Const PSt$ = "Praxisstruktur"
 Dim LizDatK$, LizDat$, LizFold$, LizFil As File, unrs$, unr&, fdt As Date
 On Error GoTo fehler
 LizFold = LokTM$ & "Lizenz\"
 LizDatK = LizFold & PSt
' If FSO Is Nothing Then Set FSO = New FileSystemObject
  LizDat = LizDatK & ".xml" '"BINTAB.DAT"
  If LenB(Dir$(LizDat)) = 0 Then GoTo Kopiere
  
' If LizFil.DateLastModified <> #4/26/2010 9:22:59 AM# Then GoTo Kopiere '#5/26/2004 8:09:00 AM# Then GoTo Kopiere
  fdt = FileDateTime(LizDat)
  If fdt = 0 Then
    Set LizFil = FSO.GetFile(LizDat)
    fdt = LizFil.DateLastModified
  End If
  Select Case fdt
   Case #4/26/2010 9:22:59 PM# '#6/17/2013 1:50:38 PM# ' weiß nicht, wo die herkommt, #4/26/2010 9:22:59 AM#
    Exit Function
   Case #4/26/2010 8:22:59 PM# '#6/17/2013 1:50:38 PM# ' weiß nicht, wo die herkommt, #4/26/2010 9:22:59 AM#
    Exit Function
  End Select
  Set LizFil = Nothing
  Do
   unr = unr + 1
   unrs = unr
'   Name LizDat As LizDatK & unrs & ".xml"
'   SuSh "cmd /c ren " & Chr$(34) & LizDat & Chr$(34) & " " & Chr$(34) & PSt & unrs & ".xml" & Chr$(34), 3
    rufauf "cmd", "/c ren """ & LizDat & """ """ & PSt & unrs & ".xml""", , , 0, 1
   If Len(Dir$(LizDat)) = 0 Then Exit Do
  Loop
Kopiere:
 On Error Resume Next
 If Not FSO.FolderExists(LizFold) Then
  FSO.CreateFolder (LizFold)
 End If
 KopDat "\\linux1\turbomed\Lizenz\" & PSt & ".xml", LizDat '"\\linux1\turbomed\BINTAB.DAT", LokTM + "BINTAB.DAT"
 Err.Clear
 Set LizFil = FSO.GetFile(LizDat)
' If Err.Number <> 0 Or (LizFil.DateLastModified <> #6/17/2013 1:50:38 PM# And LizFil.DateLastModified <> #4/26/2010 9:22:59 PM#) Then '#5/26/2004 8:09:00 AM# Then
 If Err.Number <> 0 Or (LizFil.DateLastModified <> #4/26/2010 9:22:59 PM#) Then '#5/26/2004 8:09:00 AM# Then
   KopDat "\\ANMELDR1\turbomed\Lizenz\" & PSt & ".xml", LizDat ' "\\ANMELDR1\turbomed\BINTAB.DAT", LokTM + "BINTAB.DAT"
   Err.Clear
   Set LizFil = FSO.GetFile(LizDat)
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in LizKop/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LizKop

'Function GlobalIniGet(Cpt$)
'  Dim ini$, Text$, merk%, m1%, d1%, d2, vorsp$, getDokPfad$, Eintr$(), EintrZ&, i&, fertig%
'  On Error GoTo fehler
'  ini = getReg(2, "Software\TurboMed EDV GmbH\TurboMed\Current", "RegisterPath")
'  If ini = "" Then ini = "C:\Turbomed\Programm"
'  ini = ini + IIf(Right(ini, 1) = "\", "", "\") + "global.ini"
'  merk = 0
'  m1 = 0
'  'Dim FSO As New FileSystemObject
'  If fileexists(ini) Then
'   Open ini For Input Access Read As #10
'   Do While Not EOF(10)
'    Line Input #10, Text
'    ReDim Preserve Eintr(EintrZ)
'    Eintr(EintrZ) = Text
'    EintrZ = EintrZ + 1
'   Loop
'   Close #10
'   Open ini + "1" For Output As #11
'   For i = 0 To EintrZ - 1
'    Const ADSBP$ = "Automatische Datensicherung bei Programmende?="
'    Const ASBP$ = "Automatische Datenspiegelung bei Programmende?="
'    Const ASDBP$ = "Automatische Sicherung der Dokumente bei Programmende?="
'    Const DSiBez$ = "Zielverzeichnis für Sicherung="
'    If InStr(Eintr(i), ADSBP) > 0 Then
'      Print #11, ADSBP + "{" + IIf(Cpt = "ANMELDR", "ja", "nein") + "}"
'    ElseIf InStr(Eintr(i), ASBP) > 0 Then
'      Print #11, ASBP + "{nein}"
'    ElseIf InStr(Eintr(i), ASDBP) > 0 Then
'      Print #11, ASDBP + "{nein}"
'    ElseIf InStr(Eintr(i), DSiBez) > 0 Then
'      Print #11, DSiBez + "{" + DSi + "}" ' E:\Turbomed-DASI, \\ANMELDR\TM-Dasi
'    Else
'      Print #11, Eintr(i)
'    End If
'   Next
'   Close #11
'   On Error Resume Next
'   Call Kill(ini + "alt4")
'   Call FSO.MoveFile(ini + "alt3", ini + "alt4")
'   Call FSO.MoveFile(ini + "alt2", ini + "alt3")
'   Call FSO.MoveFile(ini + "alt1", ini + "alt2")
'   Call FSO.MoveFile(ini + "alt", ini + "alt1")
''   On Error GoTo fehler
'   Call FSO.MoveFile(ini, ini + "alt")
'   Call FSO.MoveFile(ini + "1", ini)
'  End If
'  Exit Function
'fehler:
'Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in globalIniGet/" + App.Path)
' Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Function ' globalini


'Existiert in einer Registry-Section
'"HKEY_LOCAL_MACHINESOFTWAREMicrosoftOffice 10.0 / 9.0 / 8.0" der Pfad:
'"CommonInstallRoot"?
'Für Office XP existiert er nur in der Section 10.0
'Bei Office 2000 existiert er nur in 9.0
'Bei Office 97 existiert er nur in 8.0

Public Function OfficeVersion$()
 Dim RetHandle&, i%
 On Error GoTo fehler
 OfficeVersion = ""
 For i = 12 To 5 Step -1
  If RegistryPfadVorhanden(HKEY_LOCAL_MACHINE, _
  "Software\Microsoft\Office\" + CStr(i) + ".0\Common\InstallRoot") = True Then
   OfficeVersion = CStr(i) + ".0"
   Exit For
  End If
 Next
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in OfficeVersion/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' OfficeVersion

'Funktion, die für einen Vorhandenen Registry-Pfad "True", ansonsten "False", zurückgibt

Private Function RegistryPfadVorhanden(hKey As Long, Path As String) As Boolean
 Dim RetHandle&
 On Error GoTo fehler
 If RegOpenKeyEx(hKey, Path, 0&, &H1, RetHandle) = 0 Then
  RegCloseKey RetHandle
  RegistryPfadVorhanden = True
 Else
  RegistryPfadVorhanden = False
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegistryPfadVorhanden/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' RegistryPfadVorhanden
Public Sub SetMenuUnderlines(ByVal lngMenuUnderlines As Integer, Optional ByVal boolUpdateINIFile As Boolean = False)

        Dim lngUpdateINIFile%
        
        On Error Resume Next
        If boolUpdateINIFile Then lngUpdateINIFile = SPIF_UPDATEINIFILE

        Call SystemParametersInfo(SPI_SETMENUUNDERLINES, 0, lngMenuUnderlines, SPIF_SENDWININICHANGE Or lngUpdateINIFile)
        Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SetMenuUnderlines/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
    End Sub ' SetMenuUnderLines
    
#If obmitWMi Then
Function Detail_Ansicht()
 Dim binaer(), Key$, Result&
'* Script zur Änderung der Ordneransicht in W2k oder WXP/W2k3 von Norbert Fehlauer
'* Unter Verwendung von "Scripting für Adminstratoren" von Tobias Weltner
'* DeleteRegistryKey Sub von Torgeir Bakken (MVP)
Dim objWMIService As SWbemObjectSet
Dim WMIreg As SWbemObjectEx
Dim System As SWbemObjectEx
'*Konstanten deklarieren
Const HKCU = &H80000001
On Error GoTo fehler
'* Betriebssystemversion auslesen
Set objWMIService = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")

For Each System In objWMIService
'* Für Windows XP oder neuer
 Set WMIreg = GetObject("winmgmts:root\default:StdRegProv")
 Key = "Software\Microsoft\Windows\CurrentVersion\Explorer\Streams"
 Result = WMIreg.CreateKey(HKCU, Key)
 If System.Version >= "5.1.2600" Then
        binaer = Array(8, 0, 0, 0, 4, 0, 0, 0, 1, 0, 0, 0, 0, 119, 126, 19, 115, _
        53, 207, 17, 174, 105, 8, 0, 43, 46, 18, 98, 4, 0, 0, 0, 2, 0, 0, 0, 67, 0, 0, 0)
        Result = WMIreg.setbinaryvalue(HKCU, Key, "Settings", binaer)
        Key = "Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Defaults"
        Result = WMIreg.CreateKey(HKCU, Key)
        binaer = Array(28, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 144, 0, 0, 0, _
        0, 0, 1, 0, 0, 0, 255, 255, 255, 255, 240, 240, 240, 240, 20, 0, 3, 0, 144, 0, _
        0, 0, 0, 0, 0, 0, 48, 0, 0, 0, 253, 223, 223, 253, 15, 0, 4, 0, 32, 0, 16, 0, 40, _
        0, 60, 0, 0, 0, 0, 0, 1, 0, 0, 0, 2, 0, 0, 0, 3, 0, 0, 0, 245, 0, 96, 0, 120, 0, _
    120, 0, 0, 0, 0, 0, 1, 0, 0, 0, 2, 0, 0, 0, 3, 0, 0, 0, 255, 255, 255, 255, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0)
        Result = WMIreg.setbinaryvalue(HKCU, Key, "{F3364BA0-65B9-11CE-A9BA-00AA004AE837}", binaer)
        Key = "Software\Microsoft\Windows\ShellNoRoam\Bags"
        Call DeleteRegistryKey(HKCU, Key, WMIreg)
 End If

 '* Für Windows 2000
 If System.Version = "5.0.2195" Then
        binaer = Array(9, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 119, 126, 19, 115, _
        53, 207, 17, 174, 105, 8, 0, 43, 46, 18, 98, 3, 0, 0, 0, 1, 0, 0, 0)
        Result = WMIreg.setbinaryvalue(HKCU, Key, "Settings", binaer)
        Key = "Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Defaults"
        Result = WMIreg.CreateKey(HKCU, Key)
        binaer = Array(28, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 108, 0, 0, 0, _
        0, 0, 1, 0, 0, 0, 255, 255, 255, 255, 240, 240, 240, 240, 20, 0, 3, 0, 108, 0, 0, _
        0, 0, 0, 0, 0, 48, 0, 0, 0, 253, 223, 223, 253, 14, 0, 4, 0, 32, 0, 16, 0, 40, _
        0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 2, 0, 0, 0, 3, 0, 0, 0, 120, 0, 96, 0, 120, 0, _
        120, 0, 0, 0, 0, 0, 1, 0, 0, 0, 2, 0, 0, 0, 3, 0, 0, 0, 255, 255, 255, 255)
        Result = WMIreg.setbinaryvalue(HKCU, Key, "{F3364BA0-65B9-11CE-A9BA-00AA004AE837}", binaer)
 End If
Next
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Detail_Ansicht/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DetailAnsicht
'*Sub zum rekursiven Löschen von HKCU\Software\Microsoft\Windows\ShellNoRoam\Bags
Sub DeleteRegistryKey(ByVal sHive, ByVal Key, WMIreg As WbemScripting.SWbemObjectEx)
  Dim aSubKeys, sSubKey, iRC, Result
  On Error Resume Next
  Result = WMIreg.Enumkey(sHive, Key, aSubKeys)
  If Result = 0 And IsArray(aSubKeys) Then
     For Each sSubKey In aSubKeys
      If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
     End If
      Call DeleteRegistryKey(sHive, Key & "\" & sSubKey, WMIreg)
    Next
  End If
  Call WMIreg.DeleteKey(sHive, Key)
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DeleteRegistryKey/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DeleteRegistryKey
#End If

' --- Code
Public Function APIErrorDescription$(ByVal ErrLastDllError&)
' Liefert die Klartextbeschreibung zu einer API Fehlernummer, die
' unter Visual Basic über Err.LastDllError ermittelt wurde.
' HINWEIS: Die API-Funktion GetLastError ist für Visual Basic tabu!
Dim sBuffer    As String  ' String für die Rückgabe des Fehlertexts
Dim lBufferLen As Long    ' Länge des reservierten Strings
On Error GoTo fehler
  ' Stringbuffer für die Rückgabe reservieren
  sBuffer = Space$(1024)
  ' Fehlernummer in einen Fehlertext wandeln
  lBufferLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_MAX_WIDTH_MASK Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, ErrLastDllError, LANG_USER_DEFAULT, sBuffer, Len(sBuffer), 0)
  If lBufferLen > 0 Then
    ' Fehler wurde identifiziert, der Fehlertext liegt vor
    APIErrorDescription = Left$(sBuffer, lBufferLen)
  Else
    ' Der Fehlertext konnte nicht ermittelt werden
    APIErrorDescription = "Unbekannter Fehler: &H" & Hex$(ErrLastDllError)
  End If
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ApiErrorDescription/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ApiErrorDescription

Function tweakui()
' Dim WMIneu
 On Error Resume Next
' Dim WMIreg As SWbemObjectEx ' %windir%\system32\wbem\wbemdisp.tlb
' If WMIreg Is Nothing Then
'   Set WMIneu = GetObject("winmgmts:root\default:StdRegProv")
'   If Err.Number <> 0 Then Exit Function
' End If
 Dim breg As New Registry
 On Error GoTo fehler
 ' Call fDWSpei(HCU, "Control Panel\Desktop", "ForegroundFlashCount", 5)
 Call fDWSpei(HCU, "Control Panel\Desktop", "ForegroundLockTimeout", &H30D40)
 Call fStSpei(HCU, "Control Panel\Desktop", "CoolSwitchColumns", "10") ' Zahl der Spalten bei Alt+Tab
 Call fStSpei(HCU, "Control Panel\Desktop", "CoolSwitchRows", "10")
 Call fStSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "Use Search Asst", "no") ' Klassische Such im Windows Explorer
' Call WMIreg.setdwordvalue(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer", "NoFileFolderConnection", 0) ' Verbindung zwischen html-Dateien und -ordnern aufheben?
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer", "NoFileFolderConnection", 0) ' Verbindung zwischen html-Dateien und -ordnern aufheben?
 breg.WriteKey Array(0, 0, 0, 0), "link", "Software\Microsoft\Windows\CurrentVersion\Explorer", HKEY_CURRENT_USER, REG_BINARY
' If WMIreg Is Nothing Then
'  Call WMIneu.setbinaryvalue(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer", "link", Array(0, 0, 0, 0)) ' "Verknüpfung zu.." anzeigen
' Else
'  Call WMIreg.setbinaryvalue(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer", "link", Array(0, 0, 0, 0)) ' "Verknüpfung zu.." anzeigen
' End If
 Call fStSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "EncryptionContextMenu", 1) ' Klassische Such im Windows Explorer
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in tweakui/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' tweakui

Function sichZon()
 On Error GoTo fehler
' Flags einstellen
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap", "AutoDetect", 0) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap", "IntranetName", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap", "ProxyBypass", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap", "UNCAsIntranet", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "Flags", 219) '
' Intranetzone definieren
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\linserv", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\linux1", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\ANMELDR", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\ANMELDR1", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\ANMELDRNEU", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\SPO", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\SPS", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\SZNNEU", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\SPN", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\SZSNEU", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\labor", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sono", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sono1", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\ANMELDL", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\fussraum", "file", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\gsnotebook", "file", 1) '
' Intranetzone einstellen
 Call fDWSpei(HCU, "Software\Microsoft\Internet Explorer\Desktop\Components", "GeneralFlags", 4) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1001", 0) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1004", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1201", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1406", 0) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1800", 0) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1804", 0) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1A00", 0) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1C00", 196608) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1E05", 196608) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2101", 1) '
 Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "CurrentLevel", 65536) '
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SichZon/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' sichZon

Function Shares()
 On Error GoTo fehler
 If Left(Cpt, 7) = "ANMELDL" Then
   Call Shareadd("\\" & Cpt, DatenAnmL, "daten", "Daten auf \\anmeldl für Sicherheitskopie von Linux1", vNS)
   Call Shareadd("\\" & Cpt, EigDatAnmL, "U", "Ausweichordner für \\linux1\daten\eigene Dateien", vNS)
   Call Shareadd("\\" & Cpt, PatDokAnmL, "P", "Ausweichordner für \\linux1\daten\patientendokumente", vNS)
'   Call Shareadd("\\" & Cpt, "d:\mysqlbackups", "mysqlbackups", "Sicherungsordner für Mysql", vNS)
'   Call ShareAdd("c:\tmserv", "TMServ", "2.Reserveordner für \\linux1\turbomed")
   If LenB(LWTrekstor) <> 0 Then
    Call Shareadd("\\" & Cpt, LWTrekstor, "F", "WD-Laufwerk, weiterer Reserveordner für alles", "")
   End If
'   Dim f1 As New Form1
'   f1.Show
'   Call f1.ShareAdd("\\" & Environ$("COMPUTERNAME"), "f:\", "vbnetdemo", "VBnet demo test share", "")
   Call Shareadd("\\" & Cpt, alData & "\gemein", "Gemein", "Reserveordner für Verschiedenes", "")
   Call Shareadd("\\" & Cpt, DownAnmL, "down", "Ausweichordner für \\linux1\daten\down", "")
   Call Shareadd("\\" & Cpt, GeraldAnmL, "T", "Ausweichordner für \\linux1\daten\shome\gerald", "", True)
   Call Shareadd("\\" & Cpt, ReadOAnmL, "T", "Ausweichordner für \\linux1\daten\shome\geraldprivat", "")
   Call Shareadd("\\" & Cpt, KothnyAnmL, "S", "Ausweichordner für \\linux1\daten\shome\kothny", "", True)
   Call Shareadd("\\" & Cpt, ReadOKAnmL, "S", "Ausweichordner für \\linux1\daten\shome\kothnyprivat", "")
   Call Shareadd("\\" & Cpt, alDasi, "TM-Dasi", "Verzeichnis für Sicherheitskopien", "")
 ElseIf Left(Cpt, 4) = "MITTE" Then
   Call Shareadd("\\" & Cpt, mitteVol & "\DAT", "daten", "zentrales Datenverzeichnis", "")
   Call Shareadd("\\" & Cpt, mitteVol & "\gemein", "gemein", "einige Ergänzungen", "")
   Call Shareadd("\\" & Cpt, mitteVol & "\DAT\shome\gerald", "gerald", "Ausweichordner für \\linux1\daten\shome\gerald", "", True)
   Call Shareadd("\\" & Cpt, mitteVol & "\DAT\shome\gerald", "geraldprivat", "Ausweichordner für \\linux1\daten\shome\gerald", "")
   Call Shareadd("\\" & Cpt, mitteVol & "\DAT\shome\kothny", "kothny", "Ausweichordner für \\linux1\daten\shome\kothny", "", True)
   Call Shareadd("\\" & Cpt, mitteVol & "\DAT\shome\kothny", "kothnyprivat", "Ausweichordner für \\linux1\daten\shome\kothny", "")
   Call Shareadd("\\" & Cpt, mitteVol & "\obsläuft", "obsläuft", "läuft steht hier, wenn's läuft", "")
   Call Shareadd("\\" & Cpt, mitteVol & "\DAT\P", "P", "Patientendaten", "")
   Call Shareadd("\\" & Cpt, mitteRoot & "\TM-Dasi", "TM-Dasi", "Datensicherungen für Turbomed", "")
   Call Shareadd("\\" & Cpt, lokalTMExeV, "Turbomed", "Turbomed Programmverzeichnis", "") ' "c:\Turbomed",
   Call Shareadd("\\" & Cpt, mitteVol & "\DAT\U", "U", "eigene Dateien", "")
   Call Shareadd("\\" & Cpt, mitteRoot & "\DAT\down", "v", "downloads", "") ' geändert 3.1.11 zur gleichmäßigeren Auslastung
'   call shareadd(tmservcptserver,"TurboMed",
 ElseIf Left(Cpt, 7) = "ANMELDR" Then
   Call Shareadd("\\" & Cpt, arBackup & "\TurboMed", "TurboMed", "Ausweichordner auf BACKUP für \\linux1\turbomed", "") ' h:\turbomed
   Call Shareadd("\\" & Cpt, arRecover & "\", "Sicherheit", "Laufwerk RECOVER für Sicherheitskopien", "")
   Call Shareadd("\\" & Cpt, arDasi, "TM-Dasi", "Verzeichnis für Sicherheitskopien", "")
   Call Shareadd("\\" & Cpt, arData & "\eigene Dateien\Anamnese\BioWin", "BioWinBackup", "Kopie der Labor-DFÜ-Dateien")  ' GetEnvir("systemdrive") & ...
'   DSi = DSiServer
   On Error Resume Next
'   If obNot = -1 Then Call Shell(arBackup + "\Programm\ptserv32.exe")
   If obNot = -1 Then Call Shell(TMStammV & "\Programm\" & "ptserv32.exe")
   On Error GoTo fehler
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in tweakui/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Shares

Function TurbomedHerricht()
' Dim erg$
 On Error GoTo fehler
 TMExeV = getTMExeV(idt)
 lokalTMExeV = idt.LokalTurbomed
' falls poetaktiv beim letzten Mal aufgerufen war:
 If LenB(Dir$(idt.TMVerz & "Umgeschalten.txt")) <> 0 Then
  Kill idt.TMVerz & "Umgeschalten.txt"
 End If
 FPos = 50
 KWn "cgm.ico", "v:\", "c:\turbomed\symbole\desktopobjekte\"
 KWn "pumpe.ico", "v:\", "c:\turbomed\symbole\desktopobjekte\BMP\"
 If Cpt = "ANMELDR" Or Cpt = "ANMELDRNEU" Then
  Dim hnot$
  TMNotV = "L:\TurboMed\"
'  hnot = dir(TMNotV & "*.*")
'  If hnot = "" Then
  If Not FSO.FolderExists(TMNotV) Then
   TMNotV = lokalTMExeV ' "c:\turbomed\"
   If Not FSO.FolderExists(TMNotV) Then TMNotV = vNS
  End If
 Else
  TMNotV = "\\ANMELDRNEU\turbomed\"
 End If
 
 TMNotPr = TMNotV & "Programm"
 ' soll wegen Fehler in Zusammenhang mit Reklame bei der Arzneimittelliste laut Roland Colberg gelöscht werden
 If FileExists(TMExeV + "\DocPortal\bin\dpsqlite.dll") Then
  Kill (TMExeV + "\DocPortal\bin\dpsqlite.dll")
 End If
 If POk Then Print #19, "5: " + CStr(Now)
' KWn "lnk_parser_cmd.exe", "\\linux1\daten\down\", userprof
 Call AnheftNachVerz("TurboMed", TMExeV, "Turbomed.exe", TMExeV)
 Call AnheftNachVerz("TurboMed Grundeinstellungen", TMExeV, "Turbomed.exe", TMExeV, "/init")
 ' Notbetrieb: Farbeinstellungen ändern, Server = "ANMELDR", StammDB=stammDB usw., Aufrufverzeichnis auf ANMELDR ändern
' Call AnheftNachVerz("TurboMed Notbetrieb", TMExeV, "Turbomed.exe", TMNot)
 If Left$(Cpt, 7) = "ANMELDR" Then
  Call AnheftNachVerz("TM Card Server", TMExeV, "TMCardServer.exe")
'  Call AnheftNachVerz("TM Datenbank Server für Notbetrieb", TMNotPr, "ptserv32.exe", TMNotPr)
 End If
 
 Call Shares
 ' fi.Stand = "3c. zwischen Shareadd"
 
 Dim verz0$, verz1$
' verz0 = Environ("systemdrive") & "\ifapwin"
 verz0 = lokalTMExeV
' erg = Dir(verz0, vbDirectory)
' If LenB(erg) <> 0 Then If (GetAttr(verz0) And vbDirectory) <> vbDirectory Then Kill verz0
' If DirExists(verz0) Then If (GetAttr(verz0) And vbDirectory) <> vbDirectory Then Kill verz0
' If LenB(erg) = 0 Then MkDir (verz0)
 If FileExists(verz0) Then Kill verz0
 If Not DirExists(verz0) Then MkDir verz0
' verz1 = verz0 & "\hier"
 verz1 = verz0 & "ifap"
' erg = Dir(verz1, vbDirectory)
' If LenB(erg) <> 0 Then If (GetAttr(verz1) And vbDirectory) <> vbDirectory Then Kill verz1
' If LenB(erg) = 0 Then MkDir (verz1)
 If FileExists(verz1) Then Kill verz1
 If Not DirExists(verz1) Then MkDir verz1
' Call Shareadd("\\" & Cpt, verz0, "ifapwin", "Ifap Arzneimittelindex auf " & Cpt, "" & AdminGes)
' Call Shareadd("\\" & Cpt, verz1, "ifapwin", "Ifap Arzneimittelindex auf " & Cpt, "" & AdminGes)
 Call Shareadd("\\" & Cpt, lokalTMExeV, "Turbomed", "Turbomed auf " & Cpt, "" & AdminGes)
' KWn "Praxisstruktur.xml", TMStammV & "\Lizenz", lokalTMExeV & "\lizenz"
' Dim f1 As New Form1
' Call f1.ShareAdd("\\" & Cpt, verz0, "ifapwin", "Ifap Arzneimittelindex auf " & Cpt, "")
 ' fi.Stand = "4. nach Shareadd von Ifap"

 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in TurbomedHerricht/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' TurbomedHerricht

Function LWSuch$(dname$)
Const DRIVE_CDROM = 5
Const DRIVE_FIXED = 3
Const DRIVE_RAMDISK = 6
Const DRIVE_REMOTE = 4
Const DRIVE_REMOVABLE = 2
Const DRIVE_NO_ROOT_DIR = 1
Const DRIVE_UNKNOWN = 0
LWSuch = ""
 Dim sDrives$, eDr$, erg&, i%, x&, dTyp&, dBez$, dVol$
 erg = GetLogicalDriveStrings(0, sDrives)
 sDrives = String(erg, 0)
 erg = GetLogicalDriveStrings(erg, sDrives)
 For i = 1 To InStr(1, sDrives, vbNullChar & vbNullChar) Step 4
  eDr = Mid(sDrives, i, 2)
  dTyp = GetDriveType(eDr)
    Select Case dTyp
     Case DRIVE_UNKNOWN:   dBez = "unbekannt"
     Case DRIVE_NO_ROOT_DIR: dBez = "kein Wurzelverzeichnis"
     Case DRIVE_CDROM:     dBez = "CD-ROM"
     Case DRIVE_FIXED:     dBez = "Festplatte"
      dVol = Dir(eDr, vbVolume)
      If UCase$(dVol) = UCase$(dname) Then
       LWSuch = eDr
       Exit Function
      End If
     Case DRIVE_RAMDISK:   dBez = "RAM-Disk"
     Case DRIVE_REMOTE:    dBez = "Netzlaufwerk"
     Case DRIVE_REMOVABLE: dBez = "Wechseldatenträger"
    End Select
 Next i
End Function ' LWListe

Public Function getHAPDF$()
   Const FileName$ = "KVB_Arztsuche*.pdf"
   Dim jFil$, jDat As Date, erg$
   erg = Dir(vVerz & FileName)
   Do While LenB(erg) <> 0
    If erg <> "." And erg <> ".." Then
     If FileDateTime(vVerz & erg) > jDat Then
      jDat = FileDateTime(vVerz & erg)
      jFil = erg
     End If
    End If
    erg = Dir
   Loop
   getHAPDF = vVerz & jFil
End Function ' getHAPDF

Public Function ProgEnde()
'  oEnvSystem.Environment("NVerb") = ""
  Call fDWSpei(RegOrt, "SOFTWARE\GSProducts", "NVerb", 0)
  End
End Function ' Ende


'-- Ping a string representation of an IP address.
' -- Return a reply.
' -- Return long code.
Public Function ping(sAddress As String, Reply As ICMP_ECHO_REPLY) As Long
Dim hIcmp As Long
Dim lAddress As Long
Dim lTimeOut As Long
Dim StringToSend As String
'Short string of data to send
StringToSend = "hello"
'ICMP (ping) timeout
lTimeOut = 1000 'ms
'Convert string address to a long representation.
lAddress = inet_addr(sAddress)
If (lAddress <> -1) And (lAddress <> 0) Then
    'Create the handle for ICMP requests.
    hIcmp = IcmpCreateFile()
    If hIcmp Then
        'Ping the destination IP address.
        Call IcmpSendEcho(hIcmp, lAddress, StringToSend, Len(StringToSend), 0, Reply, Len(Reply), lTimeOut)
        'Reply status
        ping = Reply.Status
        'Close the Icmp handle.
        IcmpCloseHandle hIcmp
    Else
        Debug.Print "failure opening icmp handle."
        ping = -1
    End If
Else ' (lAddress <> -1) And (lAddress <> 0) Then
    ping = -1
End If ' (lAddress <> -1) And (lAddress <> 0) Then else
End Function ' ping

'Clean up the sockets.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Public Sub SocketsCleanup()
   WSACleanup
End Sub ' SockesCleanup

'Get the sockets ready.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Public Function SocketsInitialize() As Boolean
   Dim WSAD As WSADATA
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = ICMP_SUCCESS
End Function ' SocketsInitialize

'Convert the ping response to a message that you can read easily from constants.
'For more information about these constants, visit the following Microsoft Web site:
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_pingstatus.asp

Public Function EvaluatePingResponse(PingResponse As Long) As String
  Select Case PingResponse
  'Success
  Case ICMP_SUCCESS: EvaluatePingResponse = "Success!"
  'Some error occurred
  Case ICMP_STATUS_BUFFER_TO_SMALL:    EvaluatePingResponse = "Buffer Too Small"
  Case ICMP_STATUS_DESTINATION_NET_UNREACH: EvaluatePingResponse = "Destination Net Unreachable"
  Case ICMP_STATUS_DESTINATION_HOST_UNREACH: EvaluatePingResponse = "Destination Host Unreachable"
  Case ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH: EvaluatePingResponse = "Destination Protocol Unreachable"
  Case ICMP_STATUS_DESTINATION_PORT_UNREACH: EvaluatePingResponse = "Destination Port Unreachable"
  Case ICMP_STATUS_NO_RESOURCE: EvaluatePingResponse = "No Resources"
  Case ICMP_STATUS_BAD_OPTION: EvaluatePingResponse = "Bad Option"
  Case ICMP_STATUS_HARDWARE_ERROR: EvaluatePingResponse = "Hardware Error"
  Case ICMP_STATUS_LARGE_PACKET: EvaluatePingResponse = "Packet Too Big"
  Case ICMP_STATUS_REQUEST_TIMED_OUT: EvaluatePingResponse = "Request Timed Out"
  Case ICMP_STATUS_BAD_REQUEST: EvaluatePingResponse = "Bad Request"
  Case ICMP_STATUS_BAD_ROUTE: EvaluatePingResponse = "Bad Route"
  Case ICMP_STATUS_TTL_EXPIRED_TRANSIT: EvaluatePingResponse = "TimeToLive Expired Transit"
  Case ICMP_STATUS_TTL_EXPIRED_REASSEMBLY: EvaluatePingResponse = "TimeToLive Expired Reassembly"
  Case ICMP_STATUS_PARAMETER: EvaluatePingResponse = "Parameter Problem"
  Case ICMP_STATUS_SOURCE_QUENCH: EvaluatePingResponse = "Source Quench"
  Case ICMP_STATUS_OPTION_TOO_BIG: EvaluatePingResponse = "Option Too Big"
  Case ICMP_STATUS_BAD_DESTINATION: EvaluatePingResponse = "Bad Destination"
  Case ICMP_STATUS_NEGOTIATING_IPSEC: EvaluatePingResponse = "Negotiating IPSEC"
  Case ICMP_STATUS_GENERAL_FAILURE: EvaluatePingResponse = "General Failure"
  'Unknown error occurred
  Case Else: EvaluatePingResponse = "Unknown Response"
  End Select
End Function ' EvaluatePingResponse

Function doPing&(ByVal Addr$)
   Dim Reply As ICMP_ECHO_REPLY
   'Get the sockets ready.
   If SocketsInitialize() Then
   'Ping the IP that is passing the address and get a reply.
   doPing = ping(Addr, Reply)
    'Display the results.
'    Debug.Print "Address to Ping: " & Addr
'    Debug.Print "Raw ICMP code: " & lngSuccess
'    Debug.Print "Ping Response Message : " & EvaluatePingResponse(lngSuccess)
'    Debug.Print "Time : " & Reply.RoundTripTime & " ms"
    'Clean up the sockets.
    SocketsCleanup
   Else
    'Winsock error failure, initializing the sockets.
    Debug.Print WINSOCK_ERROR
    doPing = -1000
   End If
End Function ' doPing
