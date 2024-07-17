VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form formberanda 
   ClientHeight    =   6855
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   1588
      ButtonWidth     =   2355
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "BERANDA"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "PEMASUKAN"
            Key             =   ""
            Object.ToolTipText     =   "masukkan pemasukan keuangan anda disini"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "PENGELUARAN"
            Key             =   ""
            Object.ToolTipText     =   "masukkan Pengeluaran Keuangan anda"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "KELUAR"
            Key             =   ""
            Object.ToolTipText     =   "Jalan keluar ada disini ^o^"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   1508
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Bevel           =   0
            Text            =   ""
            TextSave        =   "22:27"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   10292
            Text            =   "Pelacak penggunaan Keuangan"
            TextSave        =   "Pelacak penggunaan Keuangan"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Pelacak penggunaan Keuangan"
            TextSave        =   "17/07/2024"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form beranda.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form beranda.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form beranda.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form beranda.frx":094E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Lblselisih 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Sisa Saldo Anda :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Selamat Datang"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3975
   End
End
Attribute VB_Name = "formberanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
HitungSelisihSaldo
End Sub
Private Sub HitungSelisihSaldo()
    Dim TotalSaldoPemasukan As Double
    Dim totalSaldoPengeluaran As Double
    Dim selisihSaldo As Double
    
    ' Ambil total saldo pemasukan dari FormPemasukan
    TotalSaldoPemasukan = CDbl(formpemasukan.Label5.Caption)
    
    ' Ambil total saldo pengeluaran dari FormPengeluaran
    totalSaldoPengeluaran = CDbl(Formpengeluaran.Label2.Caption)
    
    ' Hitung selisih saldo (pemasukan - pengeluaran)
    selisihSaldo = TotalSaldoPemasukan - totalSaldoPengeluaran
    
    ' Tampilkan selisih saldo dalam LabelSelisih
    Lblselisih.Caption = Format(selisihSaldo, "Currency")
End Sub
Private Sub Form_Activate()
    HitungSelisihSaldo
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
If Button.Index = 2 Then
formpemasukan.Show
End If

If Button.Index = 3 Then
Formpengeluaran.Show
End If

If Button.Index = 4 Then
End
End If
End Sub
