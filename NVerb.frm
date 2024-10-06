VERSION 5.00
Begin VB.Form FürIcon 
   Caption         =   "Startprogramm Praxis Schade"
   ClientHeight    =   4815
   ClientLeft      =   10320
   ClientTop       =   9660
   ClientWidth     =   8415
   Icon            =   "NVerb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8415
   Begin VB.Label Stand 
      Caption         =   "Stand:"
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   8295
   End
   Begin VB.Label Laufzeit 
      Caption         =   "Laufzeit:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Dies ist Ihr Computerstartprogramm. Bitte normalerweise nicht abbrechen, Sie können nebenher evtl. schon arbeiten!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "FürIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Enum enSW
    SW_HIDE
    SW_SHOWNORMAL
    SW_SHOWMINIMIZED
    SW_MAXIMIZE
    SW_SHOWNOACTIVATE
    SW_SHOW
    SW_MINIMIZE
    SW_SHOWMINNOACTIVE
    SW_SHOWNA
    SW_RESTORE
    SW_SHOWDEFAULT
    SW_FORCEMINIMIZE
End Enum



Public Function fuehraus(Datei$, Para$, Optional vz$)
Dim hierdir$
If hierdir = vbNullString Then hierdir = Environ("userprofile")
If vz <> vbNullString Then hierdir = vz
ShellExecute Me.hwnd, vbNullString, Datei, Para, hierdir, SW_HIDE
End Function

