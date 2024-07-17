VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form formloading 
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   360
      Top             =   360
   End
   Begin ComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   4440
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   105
   End
   Begin VB.Label lblpersen 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading . . ."
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "VER 1.0.0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PELACAK PENGGUNAAN KEUANGAN"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4440
      TabIndex        =   1
      Top             =   1920
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SAVER PLAN"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   0
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   480
      Picture         =   "Fromloading.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "formloading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
pb1.Value = pb1.Value + 5
lblpersen.Caption = pb1.Value & "%"
If (pb1.Value = pb1.Max) Then
Timer1.Enabled = False
Unload Me
formberanda.Show
End If
End Sub
