VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About"
   ClientHeight    =   3915
   ClientLeft      =   6375
   ClientTop       =   3000
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5565
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   $"frmAbout.frx":0000
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout.frx":009A
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0136
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

Unload Me
BrowserMain.Show

End Sub

Private Sub Form_Terminate()

Unload Me
BrowserMain.Show

End Sub
