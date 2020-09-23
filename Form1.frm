VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Flash Player"
   ClientHeight    =   2925
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4020
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _cx             =   7011
      _cy             =   5106
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu load 
         Caption         =   "&Load"
      End
      Begin VB.Menu url 
         Caption         =   "&URL"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu cont 
      Caption         =   "Co&ntrols"
      Begin VB.Menu play 
         Caption         =   "&Play"
      End
      Begin VB.Menu pause 
         Caption         =   "&Pause"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Resize()
ShockwaveFlash1.Width = Me.Width - 100
ShockwaveFlash1.Height = Me.Height - 100

End Sub

Private Sub load_Click()
CommonDialog1.ShowOpen
ShockwaveFlash1.Movie = CommonDialog1.FileName

End Sub

Private Sub pause_Click()
ShockwaveFlash1.stop
End Sub

Private Sub play_Click()
ShockwaveFlash1.play
End Sub
Private Sub url_Click()
Dim url As String
url = InputBox("Insert URL of the Flash Movie", "Play Url", "http://61.156.28.24/flash/swf/m2096.swf")
If url = "" Then
MsgBox "Bad URL", vbCritical, "ERROR"
Exit Sub
Else
ShockwaveFlash1.Movie = url

End If

End Sub
